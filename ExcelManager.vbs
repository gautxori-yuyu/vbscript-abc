' =============================================
' ARQUITECTURA MEJORADA DE GESTIÓN DE EXCEL
' =============================================
Const xlOpenXMLWorkbook = 51
Const xlOpenXMLWorkbookMacroEnabled = 52

' =============================================
' 1. SINGLETON PARA LA APLICACIÓN EXCEL (VBScript)
' =============================================

' Variable global para el singleton (fuera de la clase)
Dim g_ExcelAppInstance : Set g_ExcelAppInstance = Nothing

Class cExcelApplication
    Private m_ExcelApp
    Private m_WasRunning
    Private m_FileManager
    
    Private Sub Class_Initialize()
        Set m_FileManager = New cExcelFMFileManager
    End Sub
    
    Private Sub Class_Terminate ()
        ' Cierre defensivo (no confiable, pero útil)
    	Shutdown
    End Sub

    Public Sub InitializeExcel()
        On Error Resume Next
        Set m_ExcelApp = GetObject(, "Excel.Application")
        If Err.Number <> 0 Then
            Err.Clear()
            Set m_ExcelApp = CreateObject("Excel.Application")
            m_WasRunning = False
        Else
            m_WasRunning = m_ExcelApp.Visible
        End If
        On Error GoTo 0
        
        If Not m_ExcelApp Is Nothing Then
            m_ExcelApp.Visible = True
            m_ExcelApp.ScreenUpdating = True
            m_FileManager.SetExcelApp m_ExcelApp
        End If
    End Sub
    
    ' =============================================
    ' PROPIEDADES PÚBLICAS
    ' =============================================
    Public Property Get Application()
        If m_ExcelApp Is Nothing Then InitializeExcel()
        Set Application = m_ExcelApp
    End Property
    
    ' =============================================
    ' MÉTODOS SIMPLIFICADOS (antigua cExcelFactory)
    ' =============================================
    Public Function OpenFile(filePath, readOnly, hideWindow)
        If IsEmpty(readOnly) Then readOnly = False
        If IsEmpty(hideWindow) Then hideWindow = False
        Set OpenFile = m_FileManager.OpenFile(filePath, readOnly, hideWindow)
    End Function
    
    Public Sub CloseFile(filePath, saveChanges)
        If IsEmpty(saveChanges) Then saveChanges = False
        m_FileManager.CloseFile filePath, saveChanges
    End Sub
    
    Public Function SaveFileAs(oldPath, newPath)
        SaveFileAs = m_FileManager.SaveFileAs(oldPath, newPath)
    End Function
    
    Public Function IsFileOpenExternally(filePath)
        IsFileOpenExternally = m_FileManager.IsFileOpenExternally(filePath)
    End Function
    
    ' =============================================
    ' ACCESO SEGURO AL WORKBOOK
    ' =============================================
	' Obtiene un workbook de forma segura, validando su existencia
	' @param filePath - Ruta del archivo
	' @return Workbook válido o Nothing si no está disponible
	Public Function SafeGetWorkbook(filePath)
	    On Error Resume Next
	    Dim oExcelFMFile
	    Set SafeGetWorkbook = Nothing
	    
	    ' Intentar obtener el archivo gestionado
	    Set oExcelFMFile = m_FileManager.GetManagedFile(filePath)
	    If Not (oExcelFMFile Is Nothing) Then
	        If oExcelFMFile.IsValid Then
	            Set SafeGetWorkbook = oExcelFMFile.Workbook
	        Else
	            ' Archivo cerrado externamente, limpiarlo
	            m_FileManager.UnregisterFile filePath
	        End If
	    End If
	    
	    On Error GoTo 0
	End Function
    
    ' =============================================
    ' MÉTODO PARA RENOMBRAR ARCHIVOS
    ' =============================================
    Public Function RenameWorkbook(oldPath, newPath)
        RenameWorkbook = False
        
        ' Verificar si está abierto externamente o gestionado
        Dim oExcelFMFile
        Set oExcelFMFile = m_FileManager.GetManagedFile(oldPath)
        
        If Not (oExcelFMFile Is Nothing) Then
            ' Si está gestionado, usar SaveAs
            RenameWorkbook = SaveFileAs(oldPath, newPath)
        ElseIf IsFileOpenExternally(oldPath) Then
            ' Si está abierto externamente, necesitamos abrirlo primero
            Set oExcelFMFile = m_FileManager.OpenFile(oldPath, False, False)
            If Not (oExcelFMFile Is Nothing) Then
                RenameWorkbook = SaveFileAs(oldPath, newPath)
            End If
        Else
            ' Si no está abierto, mover archivo directamente
            On Error Resume Next
            fso.MoveFile oldPath, newPath
            RenameWorkbook = (Err.Number = 0)
            On Error GoTo 0
        End If
    End Function
    
    ' =============================================
    ' MÉTODOS DE MANIPULACIÓN DE CONTENIDO UNIFICADOS
    ' (Evitan exponer cExcelFMFile a clases externas)
    ' =============================================
	Public Function HasWorksheet(filePath, sheetName)
	    Dim oExcelFMFile
	    Set oExcelFMFile = m_FileManager.GetManagedFile(filePath)
	    If Not (oExcelFMFile Is Nothing) Then
	        HasWorksheet = oExcelFMFile.HasWorksheet(sheetName)
	    Else
	        ' comprobar en cerrado
	        Dim dicSheets
	        Set dicSheets = m_FileManager.GetSheetNamesFromClosedFile(filePath)
	        HasWorksheet = dicSheets.Exists(sheetName)
	    End If
	End Function

	' =====================================================
	' Devuelve un Array de nombres de hojas en un libro,
	' esté abierto o cerrado.
	' =====================================================
	Public Function GetSheetNames(filePath)
	    Dim oExcelFMFile, dicSheets, colSheets, ws, wsName
	    
	    Set oExcelFMFile = m_FileManager.GetManagedFile(filePath)
	    Set colSheets = CreateObject("Scripting.Dictionary")
	    
	    If Not (oExcelFMFile Is Nothing) Then
	        ' Caso abierto: recorrer las hojas
	        For Each ws In oExcelFMFile.Workbook.Worksheets
	            colSheets(ws.Name) = True
	        Next
	    Else
	        ' Caso cerrado: usar ADO
	        Set dicSheets = m_FileManager.GetSheetNamesFromClosedFile(filePath)
	        For Each wsName In dicSheets.Keys
	            colSheets(wsName) = True
	        Next
	    End If
	    
	    Set GetSheetNames = colSheets.Keys
	End Function
	
	
	' =====================================================
	' Indica si un libro (abierto o cerrado) tiene al menos
	' una hoja de trabajo.
	' =====================================================
	Public Function HasAnyWorksheet(filePath)
	    HasAnyWorksheet = (UBound(GetSheetNames(filePath)) >= 0)
	End Function
	
	
	' =====================================================
	' Devuelve el número de hojas de un libro,
	' esté abierto o cerrado.
	' =====================================================
	Public Function CountWorksheets(filePath)
	    CountWorksheets = UBound(GetSheetNames(filePath)) + 1
	End Function
	
	' =============================================
    ' GESTIÓN DEL CICLO DE VIDA
    ' =============================================
    Public Sub SetForceCloseAll(value)
        m_FileManager.ForceCloseAll = value
    End Sub
    
    Public Sub Shutdown()
        If Not m_FileManager Is Nothing Then
            m_FileManager.CloseAllSessionFiles
        End If
        
        If Not m_WasRunning And Not m_ExcelApp Is Nothing Then
            m_ExcelApp.Quit
        End If
        
        Set m_ExcelApp = Nothing
        Set m_FileManager = Nothing
        Set g_ExcelAppInstance = Nothing
    End Sub
End Class

' =============================================
' FUNCIÓN SINGLETON PARA cExcelApplication
' =============================================
Function ExcelApp()
    If g_ExcelAppInstance Is Nothing Or IsEmpty(g_ExcelAppInstance) Then
        Set g_ExcelAppInstance = New cExcelApplication
        g_ExcelAppInstance.InitializeExcel()
    End If
    Set ExcelApp = g_ExcelAppInstance
End Function


' =============================================
' 2. GESTOR DE ARCHIVOS EXCEL (Factory/Repository)
' =============================================
Class cExcelFMFileManager
    Private m_ExcelApp
    Private m_OpenFiles ' Dictionary: FilePath -> cExcelFMFile
    Private m_ForceCloseAll
    
    Private Sub Class_Initialize()
        Set m_OpenFiles = CreateObject("Scripting.Dictionary")
        m_OpenFiles.CompareMode = vbTextCompare
        m_ForceCloseAll = False
    End Sub
    
    Private Sub Class_Terminate ()
        ' Cierre defensivo
        CloseAllSessionFiles
    End Sub
    
    Public Sub SetExcelApp(excelApp)
        Set m_ExcelApp = excelApp
    End Sub
    
    Public Property Let ForceCloseAll(value)
        m_ForceCloseAll = value
    End Property
    
    Public Property Get ManagedFiles()
        Set ManagedFiles = m_OpenFiles
    End Property
    
    ' =============================================
    ' MÉTODO PÚBLICO ADICIONAL PARA cExcelApplication
    ' =============================================
    Public Function GetManagedFile(filePath)
        Set GetManagedFile = Nothing
        Dim normalizedPath
        normalizedPath = fso.GetAbsolutePathName(filePath)
        
        If m_OpenFiles.Exists(normalizedPath) Then
            Dim oExcelFMFile
            Set oExcelFMFile = m_OpenFiles(normalizedPath)
            If oExcelFMFile.IsValid Then
                Set GetManagedFile = oExcelFMFile
            End If
        End If
    End Function
    
    ' =============================================
    ' FACTORY METHOD: Crear/obtener cExcelFMFile
    ' =============================================
    Public Function OpenFile(filePath, readOnly, hideWindow)
        Dim normalizedPath
        normalizedPath = fso.GetAbsolutePathName(filePath)
        
        If Not fso.FileExists(normalizedPath) Then
            Err.Raise 53, "cExcelFMFileManager", "File not found: " & normalizedPath
        End If
        
        ' Valores por defecto para parámetros opcionales en VBScript
        If IsEmpty(readOnly) Then readOnly = False
        If IsEmpty(hideWindow) Then hideWindow = False
        
        Dim oExcelFMFile
        If m_OpenFiles.Exists(normalizedPath) Then
            Set oExcelFMFile = m_OpenFiles(normalizedPath)
            If Not oExcelFMFile.IsValid Then
                ' El archivo se cerró externamente, recrear wrapper
                m_OpenFiles.Remove normalizedPath
                Set oExcelFMFile = CreateExcelFile(normalizedPath, readOnly, hideWindow)
            End If
        Else
            Set oExcelFMFile = CreateExcelFile(normalizedPath, readOnly, hideWindow)
        End If
        
        Set OpenFile = oExcelFMFile
    End Function
    
    ' =============================================
    ' MÉTODO PRIVADO: Crear nueva instancia cExcelFMFile
    ' =============================================
    Private Function CreateExcelFile(ByVal filePath, readOnly, hideWindow)
		filePath = fso.GetAbsolutePathName(filePath)
		
        Dim workbook, wasAlreadyOpen, sessionOpened
        
        ' Verificar si ya está abierto en Excel
        Set workbook = FindOpenWorkbook(filePath)
        wasAlreadyOpen = Not (workbook Is Nothing)
        
        If wasAlreadyOpen Then
            sessionOpened = False
        Else
            ' Abrir el archivo
			'FileName, UpdateLinks (0, no; 3; yes, external references), ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify, Converter, AddToMru, Local, CorruptLoad
			' format = delimitador CSV: 1  Pestañas; 2  Comas; 3  Espacios; 4  Punto y coma; 5  Nothing
			' IgnoreReadOnlyRecommended: PARA QUE NO PREGUNTE Abrir como solo lectura, si estuviera abierto...; PARECE QUE DEVUELVE LA INSTANCIA QUE ESTUVIERA ABIERTA!!
			' Origin = 2, == crLf; ¿¿3 == Lf??
			' Editable: Si el archivo es una plantilla de Excel, True para abrir la plantilla especificada para su edición. False para abrir un nuevo libro basado en la plantilla especificada. El valor predeterminado es False.
			' Notify = False (no merece poner true, ver ayuda)
			' Converter: para seleccionar el convertidor de fichero, de los que hay en Application.FileConverters (ver ayuda para mas info)
			' AddToMru = False, PARA NO DEJAR HUELLA DE QUE SE HA ABIERTO EL FICHERO
			' Local: True guarda los archivos contra el idioma de Microsoft Excel (incluida la configuración del panel de control). False (valor predeterminado) guarda los archivos contra el idioma de Visual Basic para aplicaciones (VBA) (que suele ser inglés (Estados Unidos)
			' CorruptLoad:   2  El libro se abre en el modo de extracción de datos.;   0  El libro se abre normalmente.;   1  El libro se abre en el modo de reparación
            Set workbook = m_ExcelApp.Workbooks.Open(filePath, 3, readOnly, 2, "", "", False, 2, "", True, False, , False, True, 0)
            sessionOpened = True
        End If
        
        If hideWindow And Not workbook Is Nothing Then
            m_ExcelApp.Windows(workbook.Name).Visible = False
        End If
        
        ' FACTORY: Crear wrapper cExcelFMFile
        Dim oExcelFMFile
        Set oExcelFMFile = New cExcelFMFile
        oExcelFMFile.Initialize workbook, wasAlreadyOpen, sessionOpened, Me
        
        ' Registrar en colección gestionada
        m_OpenFiles.Add filePath, oExcelFMFile
        Set CreateExcelFile = oExcelFMFile
    End Function
    
    Private Function FindOpenWorkbook(ByVal filePath)
		filePath = fso.GetAbsolutePathName(filePath)
		
        Set FindOpenWorkbook = Nothing
        Dim wb
        For Each wb In m_ExcelApp.Workbooks
            If StrComp(wb.FullName, filePath, vbTextCompare) = 0 Then
                Set FindOpenWorkbook = wb
                Exit For
            End If
        Next
    End Function
    
    ' =============================================
    ' GESTIÓN DEL CICLO DE VIDA DE ARCHIVOS
    ' =============================================
	Public Sub UnregisterFile(ByVal filePath)
		If m_OpenFiles.Exists(filePath) Then
			m_OpenFiles.Remove filePath
		End If
	End Sub
	
	Sub UpdateFileKey(oldPath, oExcelFMFile)
	    If m_OpenFiles.Exists(oldPath) Then
	        m_OpenFiles.Remove oldPath
	    End If
	    Set m_OpenFiles(oExcelFMFile.FilePath) = oExcelFMFile
	End Sub
	
    Public Sub CloseFile(filePath, saveChanges)
		On Error Resume Next  ' Añadir protección
		filePath = fso.GetAbsolutePathName(filePath)
		
        If IsEmpty(saveChanges) Then saveChanges = False
        If m_OpenFiles.Exists(filePath) Then
            Dim oExcelFMFile
            Set oExcelFMFile = m_OpenFiles(filePath)

			' Validar antes de cerrar
			If Not (oExcelFMFile Is Nothing) Then
			    If oExcelFMFile.IsValid Then
			        oExcelFMFile.Close saveChanges
			    End If
			End If
            If m_OpenFiles.Exists(filePath) Then m_OpenFiles.Remove filePath
        End If
	    If Err.Number <> 0 Then
	        ' Log silencioso del error pero continuar
	        If Not (MsgIE Is Nothing) Then
	            MsgLog "Advertencia al cerrar " & fso.GetFileName(filePath) & ": " & Err.Description
	        End If
	        Err.Clear
	    End If
	    On Error GoTo 0
    End Sub
    
    Public Function SaveFileAs(oldPath, newPath)
        If Not m_OpenFiles.Exists(oldPath) Then
            'Err.Raise 5, "cExcelFMFileManager", "File not managed: " & oldPath
            SaveFileAs = False
            Exit Function
        End If
        
        Dim oExcelFMFile,filetype
        Set oExcelFMFile = m_OpenFiles(oldPath)
		Select Case LCase(Mid(oldPath,InstrRev(oldPath,".")))
			Case ".xlsx" : filetype = 51
			Case ".xlsm" : filetype = 52
			Case Else : Stop : filetype = 51
		End Select
        If oExcelFMFile.SaveAs(newPath,filetype) Then
            ' Actualizar la clave en el diccionario
            Set m_OpenFiles(newPath) = oExcelFMFile
            m_OpenFiles.Remove oldPath
            SaveFileAs = True
        Else
            SaveFileAs = False
        End If
    End Function
    
    Public Sub CloseAllSessionFiles()
        Dim filePath, oExcelFMFile
        For Each filePath In m_OpenFiles.Keys
            Set oExcelFMFile = m_OpenFiles(filePath)
            If oExcelFMFile.SessionOpened Or m_ForceCloseAll Then
                oExcelFMFile.Close False
            End If
        Next
        m_OpenFiles.RemoveAll
    End Sub
    
    ' =============================================
    ' VERIFICACIÓN DE ESTADO DE ARCHIVOS
    ' =============================================
    
    ' Verifica si un archivo está abierto fuera de nuestra gestión
    Public Function IsFileOpenExternally(filePath)
        Dim normalizedPath
        normalizedPath = fso.GetAbsolutePathName(filePath)
        
        ' Primero verificar si lo tenemos gestionado
        If m_OpenFiles.Exists(normalizedPath) Then
            Dim oExcelFMFile
            Set oExcelFMFile = m_OpenFiles(normalizedPath)
            ' Si está gestionado pero no válido, está abierto externamente
            IsFileOpenExternally = Not oExcelFMFile.IsValid
        Else
            ' No está gestionado, verificar si está abierto externamente
            ' Método 1: Intentar abrirlo para escritura
            On Error Resume Next
            Dim testFile
            Set testFile = fso.OpenTextFile(normalizedPath, 8, False) ' Modo append
            If Err.Number = 0 Then
                testFile.Close
                IsFileOpenExternally = False ' No está abierto
            Else
                ' Método 2: Verificar en la aplicación Excel actual
                Dim wb
                Set wb = FindOpenWorkbook(normalizedPath)
                IsFileOpenExternally = Not (wb Is Nothing)
            End If
            On Error GoTo 0
        End If
    End Function
    
    
    ' =============================================
    ' GESTIÓN DE ARCHIVOS NO ABIERTOS
    ' =============================================
    
    ' Obtiene las hojas (worksheets) de un archivo no abierto
	Public Function GetSheetNamesFromClosedFile(filePath)
		' OJO, ESTA FUNCION NO VA EN 32 BITS, si ADO está en Windows de 64 bits!!; por tanto NO VA DESDE EL VBSEDIT DE 32 BITS!!
	    Dim objConnection, objRecordset, oDic, sResult
	    Set oDic = CreateObject("Scripting.Dictionary")
	    
	    ' Abrir conexión ADO al Excel cerrado
	    Set objConnection = CreateObject("ADODB.Connection")
												
	    On Error Resume Next
	    'try 16.0
	    objConnection.Open "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & filePath & "; Extended Properties=""Excel 12.0;"""
	    If Err.Number <> 0 Then
	        'try 12.0
	        Err.Clear
	        objConnection.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & filePath & "; Extended Properties=""Excel 12.0;"""
	    End If
	    On Error GoTo 0
	    
	    If objConnection.State = 1 Then ' Abierta
													   
	        Const adSchemaTables = 20
	        Set objRecordset = objConnection.OpenSchema(adSchemaTables)
	        Do Until objRecordset.EOF
	            sResult = objRecordset.Fields.Item("TABLE_NAME")
				'If there is a space in the sheet name then remove the first and last single quotes and the $ Sign
	            If InStr(sResult, " ") > 0 Then
	                sResult = Mid(sResult, 2, Len(sResult) - 3)
	            Else
					'If there is no space then we need to remove only the $ sign form the end
	                sResult = Left(sResult, Len(sResult) - 1)
	            End If
	            If Not oDic.Exists(sResult) Then
	                oDic.Add sResult, Empty
	            End If
	            objRecordset.MoveNext
	        Loop
	        objRecordset.Close
	    End If
	    
	    objConnection.Close
	    Set GetSheetNamesFromClosedFile = oDic
	End Function
End Class

' =============================================
' 3. WRAPPER DE ARCHIVO EXCEL INDIVIDUAL
' =============================================
Class cExcelFMFile
    Private m_Workbook
    Private m_WasAlreadyOpen
    Private m_SessionOpened
    Private m_FileManager
    Private m_IsValid
    
    Private Sub Class_Initialize()
    	Set m_Workbook = Nothing
    	Set m_FileManager = Nothing
    End Sub
    
    Private Sub Class_Terminate()
        ' Cierre defensivo
        If m_IsValid And Not m_WasAlreadyOpen And Not (m_Workbook Is Nothing) Then
            m_Workbook.Close False
        End If
    End Sub
    
    Public Sub Initialize(workbook, wasAlreadyOpen, sessionOpened, fileManager)
        Set m_Workbook = workbook
        m_WasAlreadyOpen = wasAlreadyOpen
        m_SessionOpened = sessionOpened
        Set m_FileManager = fileManager
        m_IsValid = True
    End Sub
    
    Public Property Get Workbook()
        If Not m_IsValid Then
            Err.Raise 91, "cExcelFMFile", "Excel file is no longer valid"
        End If
        
        ' Verificar que el workbook sigue siendo válido
        On Error Resume Next
        Dim testName
        testName = m_Workbook.Name
        If Err.Number <> 0 Then
            m_IsValid = False
            Err.Raise 91, "cExcelFMFile", "Workbook reference is no longer valid"
        End If
        On Error GoTo 0
        
        Set Workbook = m_Workbook
    End Property
    
    Public Property Get FilePath()
        FilePath = m_Workbook.FullName
    End Property
    
    Public Property Get WasAlreadyOpen()
        WasAlreadyOpen = m_WasAlreadyOpen
    End Property
    
    Public Property Get SessionOpened()
        SessionOpened = m_SessionOpened
    End Property
    
    Public Property Get IsValid()
        IsValid = m_IsValid And Not (m_Workbook Is Nothing)
    End Property
    
    Public Function SaveAs(newPath,filetype)
        If Not m_IsValid Then
            SaveAs = False
            Exit Function
        End If
        
	    Dim oldPath
	    oldPath = m_Workbook.FullName   ' clave actual ANTES de sobrescribir
	    
        On Error Resume Next
		' const xlOpenXMLWorkbook = 51 ' formato XLSX
		' const xlOpenXMLWorkbook = 52 ' formato XLSM, con macros
		' expresión. SaveAs (FileName, FileFormat, Password, WriteResPassword, ReadOnlyRecommended, CreateBackup, AccessMode, ConflictResolution, AddToMru, TextCodepage, TextVisualLayout, Local)
        m_Workbook.SaveAs newPath, filetype, Empty, Empty, Empty, False, Empty, Empty, False
	    If Err.Number = 0 Then
	        ' Notificar al FileManager para actualizar el diccionario
	        m_FileManager.UpdateFileKey oldPath, Me
	        SaveAs = True
	    Else
	        SaveAs = False
	    End If
        On Error GoTo 0
    End Function
    
    Public Sub Save()
        If m_IsValid And Not m_Workbook Is Nothing Then
            m_Workbook.Save
        End If
    End Sub
    
    Public Sub Close(saveChanges)
    	Dim strTmpPath
        If IsEmpty(saveChanges) Then saveChanges = False
        If m_IsValid And Not m_Workbook Is Nothing Then
            If Not m_WasAlreadyOpen Then
            	strTmpPath = Me.FilePath
                m_Workbook.Close saveChanges
				m_FileManager.UnregisterFile strTmpPath
            End If
            m_IsValid = False
        End If
    End Sub
    
    ' =============================================
    ' MÉTODOS DE CONVENIENCIA
    ' =============================================
    Public Function HasWorksheet(sheetName)
        HasWorksheet = False
        If Not m_IsValid Then Exit Function
        
        On Error Resume Next
        Dim ws
        Set ws = m_Workbook.Worksheets(sheetName)
        HasWorksheet = (Err.Number = 0) And Not (ws Is Nothing)
        On Error GoTo 0
    End Function
    
    Public Function GetWorksheet(sheetName)
        Set GetWorksheet = Nothing
        If Not m_IsValid Then Exit Function
        
        On Error Resume Next
        Set GetWorksheet = m_Workbook.Worksheets(sheetName)
        On Error GoTo 0
    End Function
    
    Public Sub UpdateDocumentProperties(title, subject, author, keywords, comments, Manager, Company, RevisionNumber)
        If Not m_IsValid Then Exit Sub
        
        On Error Resume Next
        With m_Workbook.BuiltinDocumentProperties
            If title <> "" Then .Item("Title") = title
            If subject <> "" Then .Item("Subject") = subject
            If author <> "" Then .Item("Author") = author
            If keywords <> "" Then .Item("Keywords") = keywords
            If comments <> "" Then .Item("Comments") = comments
            If Manager <> "" Then .Item("Manager") = Manager
            If Company <> "" Then .Item("Company") = Company
            If RevisionNumber <> "" Then .Item("Revision Number") = RevisionNumber
        End With
        On Error GoTo 0
    End Sub
End Class
