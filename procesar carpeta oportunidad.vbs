Option Explicit
Dim bLog : bLog = True : If InStr(WScript.FullName,"cscript") = 0 Then bLog = False
Const bQuickExcel = True
Dim WshShell
Set WshShell = WScript.CreateObject("Wscript.Shell")
Dim fso
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Dim regex
Set regex = New RegExp
regex.Global = True
regex.IgnoreCase = True
regex.multiline = false ' Si es true, ^ y $ detectan posiciones de comienzo y final de linea

Include "constants_globals.vbs"
Include "fUtils.vbs"
Include "cHTALogger.vbs"  ' Clase HtaLogger (usada por HTMLWindow)

Dim MsgIE
Set MsgIE = (New HTMLWindow).CreaVentanaHTML ("procesado oportunidades", True, 800, 650)', Empty, Empty)'

Dim arrFolds
arrFolds = SeleccCarpetasWSArgs ("",False,"C:\abc compressors\INTRANET\OilGas\3_OFERTAS\OFERTAS\")
If UBound(arrFolds) < 0 Then WScript.Quit (1)

' Include "cMsgIEReporter.vbs"  ' [OBSOLETO] Ya no se usa, eliminado del proyecto
Include "ExcelManager.vbs"
Include "cOportunidad.vbs"
Include "cCompressor.vbs"
Include "cOp_CalcsTecn.vbs"
Include "cABCGas.vbs"
Include "cOp_ValsEcon.vbs"
Include "cOferGas.vbs"
Include "cOp_Ofertas.vbs"

' =============================================
' PROCESAMIENTO PRINCIPAL
' =============================================
Dim arg,oOportunidad,strFName
For Each arg In arrFolds
	If bQuit Then WScript.Quit
	If fso.FolderExists (arg) Then
		strFName = fso.GetFolder(arg).Name
		MsgIE.objDocument.title = Replace (MsgIE.objDocument.title,"procesado oportunidades", strFName & " / " & "procesado oportunidades")
		Set oOportunidad = (New cOportunidad).Init(arg, ExcelApp(), MsgIE, strFName)
        Call oOportunidad.procesaCarp()
		WshShell.Popup "Finalizado el procesado de la oportunidad: " & strFName, 20
		MsgBox ("NO OLVIDES SACAR LA HOJA API")
	End If
Next

' =============================================
' LIMPIEZA FINAL (solo si se usó Excel)
' =============================================
If IsEmpty(g_ExcelAppInstance) Then
ElseIf g_ExcelAppInstance Is Nothing Then
Else
	' esto se haria automaticamente con el destructor...
    ' Cerrando sistema de gestión de Excel...
    g_ExcelAppInstance.SetForceCloseAll True
    g_ExcelAppInstance.Shutdown
    Set g_ExcelAppInstance = Nothing
End If

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub Include(fSpec)
	' Sub to load file modules (Once!)
	' Simply call this function with the path to the vbs file you want to include, and it will be 
	' loaded into memory ONCE, even if you try to load it a number of times. 
	Dim sTemp
	On Error Resume Next
	sTemp = Eval("ICLF")
	On Error Goto 0
	If IsEmpty(sTemp) Then 
		' no currently loaded files - first Include run
		ExecuteGlobal "Dim ICLF : ICLF = 0"
	End If
    
	' test to see if file has already been loaded
	If InStr(1,iclf,fspec,vbTextCompare)=0 Then
		With CreateObject("Scripting.FileSystemObject")
			If .fileexists(fspec) Then
                On Error Resume Next
                ExecuteGlobal .openTextFile(fSpec).readAll()
                If Err.Number <> 0 Then
                    MsgBox "Error loading include file " & fspec & ": " & Err.Description, vbOKOnly+vbCritical, "Include Error"
                    WScript.Quit 1
                End If
                On Error GoTo 0
            Else
                MsgBox "Include file " & fspec & " not found. Exiting", vbOKOnly+vbExclamation, "Critical Error."
                WScript.Quit 1
			End If
		End With
		ICLF = ICLF & "|" & fspec
	Else
		' file already loaded
	End If
End Sub
