<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,5000)

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


	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr bgcolor='#CFAB87'><td height='20px'><strong>"&anio&"</strong></td></tr><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")

	Set objFieldsC = rsC.Fields

	response.write("<tr bgcolor='#FFE2C6'>")
	For intLoop = 0 To (objFieldsC.Count - 1)
        response.write("<td><strong>"&objFieldsC.Item(intLoop).Name&"</strong></td>")   	
    Next
    response.write("</tr>")

	rsC.MoveFirst
	response.write("<tr>")
    while not rsC.eof
	    For intLoop = 0 To (objFieldsC.Count - 1)
	        response.write("<td align='left'>"&rsC(intLoop)&"</td>")             	
	    Next      	
    	response.write("</tr>")
		rsC.MoveNext
	wend
	

	rsC.Close
	Set rsC=Nothing
	response.write("</table></td></tr>")    

	response.write("<tr bgcolor='#CFAB87'><td height='20px'><strong>"&anio1&"</strong></td></tr><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")

	Set objFieldsC2 = rsC2.Fields

	response.write("<tr bgcolor='#FFE2C6'>")
	For intLoop = 0 To (objFieldsC2.Count - 1)
        response.write("<td><strong>"&objFieldsC2.Item(intLoop).Name&"</strong></td>")   	
    Next
    response.write("</tr>")

	rsC2.MoveFirst
	response.write("<tr>")
    while not rsC2.eof
	    For intLoop = 0 To (objFieldsC2.Count - 1)
	        response.write("<td align='left'>"&rsC2(intLoop)&"</td>")         	
	    Next      	
    	response.write("</tr>")
		rsC2.MoveNext
	wend
	
	rsC2.Close
	Set rsC2=Nothing
	response.write("</table></td></tr></table>")    
%>
