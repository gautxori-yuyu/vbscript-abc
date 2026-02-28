Option Explicit

Function procesaCarpOfertas (strFold, oDicValoracs, ExcelFactory)
	If Not fso.FolderExists (strFold) Then Exit Function
	' Usa las plantillas, renombrandolas, AÑADIENDO LOS CALCULOS GAS_VBNET, metiendo los PRECIOS, Y NUMERO DE COMPRESORES... E INCLUSO GENERANDO LAS OPCIONES!!!, desde "OTROS"
	MsgIE ("<font color=blue>cosas a tener en cuenta en los ficheros de oferta:" & vbLf & _
			"<li> EL NOMBRE DEL FICHERO DE OFERTA NO PUEDE CONTENER CARACTERES ESPECIALES, coma, etc, PARA QUE SE PUEDA SUBIR A CREATIO!!</li>" & vbLf & _
			"<li> PROCESAR OTROS FICHEROS: 8000 hour & commisioning spares, packing list, etc</li></font>")
	Dim strNumOferta,strFPath
	For Each strNumOferta In oDicValoracs
	For Each strFPath In oDicValoracs(strNumOferta).Keys
		'oDicValoracs(strNumOferta)(strFichPath)
		'Stop
	Next
	Next
	'Stop
	MsgIE ("<font color=blue><b>PROCESAR OTROS FICHEROS: 8000 hour & commisioning spares, packing list, etc</b></font>")
End Function
