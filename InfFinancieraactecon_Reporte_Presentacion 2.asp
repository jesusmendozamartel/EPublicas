<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,5000)

Response.Charset= "ISO-8859-1" 

	Dim rs, sql
	Set rs=Server.CreateObject("ADODB.RecordSet")	
	
	anio=Request.QueryString("anio")
	formato=Request.QueryString("for")
	sector=Request.QueryString("sec")

	rs.CursorLocation=3

	rs.Open "exec sp_EstadosFinancieros_ListarPivot_AnioFormatoSectorGContable '"&anio&"','"&formato&"','"&sector&"'", Con

	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")

	Set objFields = rs.Fields
		response.write("<tr bgcolor='#FFE2C6'>")
	For intLoop = 0 To (objFields.Count - 1)
        response.write("<td><strong>"&objFields.Item(intLoop).Name&"</strong></td>")        	
    Next
        response.write("</tr>")


    while not rs.eof
		response.write("<tr >")
		For intLoop = 0 To (objFields.Count - 1)
	        response.write("<td>"&rs(intLoop)&"</td>")        	
	    Next
	    response.write("</tr>")
		rs.MoveNext
	wend

	rs.Close
	Set rs=Nothing
	response.write("</table></td></tr></table>")    
%>
