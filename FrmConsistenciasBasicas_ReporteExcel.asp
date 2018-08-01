<!--#include file="Conexion.asp"-->

<%
	Response.Charset= "ISO-8859-1" 
	
	anio=Request.QueryString("anio")
	anexo=Request.QueryString("ane")
	sector=Request.QueryString("sec")

	anio1=anio-1

	Dim rsC
	Set rsC=Server.CreateObject("ADODB.RecordSet")	
	rsC.CursorLocation=3

	rsC.Open "exec sp_EstadosFinancieros_ListarAnexos_AnioSectorAnexo "&anio&",'"&sector&"','"&anexo&"',1", Con

	Dim rsC2
	Set rsC2=Server.CreateObject("ADODB.RecordSet")	
	rsC2.CursorLocation=3

	rsC2.Open "exec sp_EstadosFinancieros_ListarAnexos_AnioSectorAnexo "&anio&",'"&sector&"','"&anexo&"',0", Con

	Archivo="ConsBasicasAnexo_"&anexo&"_"&anio
	Titulo="SICON - Consistencias Basicas Anexo "&anexo&"-"&anio

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+"_"+annio+".xls" 

response.Write("MIME-Version: 1.0" &vbCr) 
response.Write("X-Document-Type: Workbook" &vbCr) 
response.Write("Content-Type: multipart/related; boundary=3D""----=_NextPart_ExcelWorkbook"""&vbCr) 
response.Write(""&vbCr) 
response.Write("------=_NextPart_ExcelWorkbook" &vbCr) 
response.Write("Content-Location: books.xls" &vbCr) 
response.Write("Content-Transfer-Encoding: quoted-printable" &vbCr) 
response.Write("Content-Type: text/html; charset=3D""us-ascii""" &vbCr) 
response.Write("" &vbCr) 
response.Write("<html xmlns:x=3D""urn:schemas-microsoft-com:office:excel"">" &vbCr) 
response.Write("<head>" &vbCr) 
response.Write("<meta name=3D""Excel Workbook Frameset"">" &vbCr) 
response.Write("<xml>" &vbCr) 
      response.Write("<x:ExcelWorkbook>" &vbCr) 
            response.Write("<x:ExcelWorksheets>" &vbCr) 
                  response.Write("<x:ExcelWorksheet>" &vbCr)
                        response.Write("<x:Name>"&anio&"</x:Name>" &vbCr) 
                        response.Write("<x:WorksheetSource HRef=3D""sheet1.htm""/>" &vbCr) 
                  response.Write("</x:ExcelWorksheet>" &vbCr) 
                  response.Write("<x:ExcelWorksheet>" &vbCr) 
                        response.Write("<x:Name>"&anio1&"</x:Name>" &vbCr) 
                        response.Write("<x:WorksheetSource HRef=3D""sheet2.htm""/>" &vbCr) 
                  response.Write("</x:ExcelWorksheet>" &vbCr) 
            response.Write("</x:ExcelWorksheets>" &vbCr) 
            response.Write("<x:Stylesheet HRef=3D""stylesheet.css""/>" &vbCr) 
      response.Write("</x:ExcelWorkbook>" &vbCr) 
response.Write("</xml>" &vbCr) 
response.Write("</head>" &vbCr) 
response.Write("</html>" &vbCr) 
response.Write(""&vbCr) 
response.Write("------=_NextPart_ExcelWorkbook"&vbCr) 
response.Write("Content-Location: stylesheet.css"&vbCr) 
response.Write("Content-Transfer-Encoding: quoted-printable"&vbCr) 
response.Write("Content-Type: text/css; charset=""us-ascii"""&vbCr) 
response.Write(""&vbCr) 
response.Write(".titulo"&vbCr) 
response.Write("{"&vbCr) 
	response.Write("font-size:50px;"&vbCr) 
	response.Write("color:#000000"&vbCr) 
response.Write("}"&vbCr) 
response.Write("TABLE"&vbCr) 
response.Write("{"&vbCr) 
    response.Write("BORDER-RIGHT: #c0c0c0 1px dotted;"&vbCr) 
    response.Write("BORDER-TOP: #c0c0c0 1px dotted ;"&vbCr) 
    response.Write("BORDER-LEFT: #c0c0c0 1px dotted ;"&vbCr) 
    response.Write("BORDER-BOTTOM: #c0c0c0 1px dotted;"&vbCr) 
    response.Write("BORDER-COLLAPSE: collapse;"&vbCr) 
    response.Write("border-spacing: 0;"&vbCr) 
	response.Write("font-family:Arial, Geneva, sans-serif;"&vbCr) 
	response.Write("font-size:10px;"&vbCr) 
	response.Write("width:100%;"&vbCr) 
	response.Write("vnd.ms-excel.numberformat:#,##0;"&vbCr) 
response.Write("}"&vbCr) 
response.Write("TABLE.TD"&vbCr) 
response.Write("{"&vbCr) 
    response.Write("BORDER-RIGHT: #828282 1px dotted ;"&vbCr) 
    response.Write("BORDER-TOP: #828282 1px dotted ;"&vbCr) 
    response.Write("BORDER-LEFT: #828282 1px dotted ;"&vbCr) 
    response.Write("BORDER-BOTTOM: #828282 1px dotted ;"&vbCr) 
	response.Write("font-family:Arial, Geneva, sans-serif;"&vbCr) 
	response.Write("font-size:10px;"&vbCr) 
response.Write("}"&vbCr) 
response.Write(""&vbCr) 
response.Write("------=_NextPart_ExcelWorkbook" &vbCr) 
response.Write("Content-Location: sheet1.htm" &vbCr)
response.Write("Content-Transfer-Encoding: quoted-printable" &vbCr)
response.Write("Content-Type: text/html; charset=3D""us-ascii""" &vbCr) 
response.Write(""&vbCr) 
response.Write("<html xmlns:v=3D""urn:schemas-microsoft-com:vml""" &vbCr) 
response.Write("xmlns:o=3D""urn:schemas-microsoft-com:office:office""" &vbCr) 
response.Write("xmlns:x=3D""urn:schemas-microsoft-com:office:excel""" &vbCr) 
response.Write("xmlns=3D""http://www.w3.org/TR/REC-html40"">" &vbCr) 
response.Write("" &vbCr) 
response.Write("<head>" &vbCr) 
response.Write("<meta http-equiv=3DContent-Type content=3D""text/html; charset=3Dus-ascii"">" &vbCr) 
response.Write("<meta name=3DProgId content=3DExcel.Sheet>" &vbCr) 
response.Write("<meta name=3DGenerator content=3D""Microsoft Excel 9"">" &vbCr) 
response.Write("<link rel=3DStylesheet href=3Dstylesheet.css>" &vbCr) 
response.Write("</head>" &vbCr) 
response.Write("" &vbCr) 
response.Write("<body>" &vbCr) 
	Response.Write("<table><tr><td colspan='14'> <h2>"&Titulo&"</h2></td></tr><tr><td>&nbsp;&nbsp;</td></tr>")			
	Set objFieldsC = rsC.Fields

	response.write("<tr bgcolor='#FFE2C6'>")
	For intLoop = 0 To (objFieldsC.Count - 1)
        response.write("<td><strong>"&objFieldsC.Item(intLoop).Name&"</strong></td>")   	
    Next
    response.write("</tr>")

	rsC.MoveFirst
    while not rsC.eof
		response.write("<tr>")
	    For intLoop = 0 To (objFieldsC.Count - 1)
	        response.write("<td align='left' style='vnd.ms-excel.numberformat:#,##0.00;'>"&rsC(intLoop)&"</td>")             	
	    Next      	
    	response.write("</tr>")
		rsC.MoveNext
	wend


	rsC.Close
	Set rsC=Nothing
	response.write("</table>")    
response.Write(""&vbCr) 
response.Write("------=_NextPart_ExcelWorkbook" &vbCr) 
response.Write("Content-Location: sheet2.htm" &vbCr) 
response.Write("Content-Transfer-Encoding: quoted-printable" &vbCr) 
response.Write("Content-Type: text/html; charset=3D""us-ascii""" &vbCr) 
response.Write(""&vbCr) 
response.Write("<html xmlns:v=3D""urn:schemas-microsoft-com:vml""" &vbCr) 
response.Write("xmlns:o=3D""urn:schemas-microsoft-com:office:office""" &vbCr) 
response.Write("xmlns:x=3D""urn:schemas-microsoft-com:office:excel""" &vbCr) 
response.Write("xmlns=3D""http://www.w3.org/TR/REC-html40"">" &vbCr) 
response.Write("" &vbCr) 
response.Write("<head>" &vbCr) 
response.Write("<meta http-equiv=3DContent-Type content=3D""text/html; charset=3Dus-ascii"">" &vbCr) 
response.Write("<meta name=3DProgId content=3DExcel.Sheet>" &vbCr) 
response.Write("<meta name=3DGenerator content=3D""Microsoft Excel 9"">" &vbCr) 
response.Write("<link rel=3DStylesheet href=3Dstylesheet.css>" &vbCr) 
response.Write("</head>" &vbCr) 
response.Write("" &vbCr) 
response.Write("<body>" &vbCr) 
	Response.Write("<table><tr><td colspan='14'> <h2>"&Titulo&"</h2></td></tr><tr><td>&nbsp;&nbsp;</td></tr>")			
	Set objFieldsC = rsC2.Fields

	response.write("<tr bgcolor='#FFE2C6'>")
	For intLoop = 0 To (objFieldsC.Count - 1)
        response.write("<td><strong>"&objFieldsC.Item(intLoop).Name&"</strong></td>")   	
    Next
    response.write("</tr>")

	rsC2.MoveFirst
    while not rsC2.eof
		response.write("<tr>")
	    For intLoop = 0 To (objFieldsC.Count - 1)
	        response.write("<td align='left' style='vnd.ms-excel.numberformat:#,##0.00;'>"&rsC2(intLoop)&"</td>")             	
	    Next      	
    	response.write("</tr>")
		rsC2.MoveNext
	wend

	rsC2.Close
	Set rsC2=Nothing
	response.write("</table>")
response.Write("</body>" &vbCr) 
response.Write("</html>" &vbCr) 
response.Write("------=_NextPart_ExcelWorkbook--" &vbCr) 
%>

