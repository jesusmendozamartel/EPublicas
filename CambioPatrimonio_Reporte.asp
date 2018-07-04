<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,2000)
Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	nivel=Request.QueryString("nivel")
	codigo=Request.QueryString("codigo")
	'SecSic=Request.QueryString("secsic")
	gcont=Request.QueryString("gcont")
	detalle=Request.QueryString("detalle")

	SQL="EXEC sp_lista_directorio_RepAnioGContNiv '04','"&annio&"','"&nivel&"','"&codigo&"','"&gcont&"','"&detalle&"'"

	Set rs = Server.CreateObject("ADODB.Recordset")

	rs.CursorLocation=3
	rs.Open SQL, con 

	Response.Write("<tr class='a1'> RECUENTO: "&rs.RecordCount&"</tr>")

	x=rs.Fields.Count-1
	
	j=0

	Response.Write("<table class='tabla1'>")
	Response.Write("<th >Ver Detalle</th>")

	for i=0 to x 
		Response.Write("<th >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		Response.Write("<td  align="&alig&"><button onClick=""cargaDetCamPat('"&Rs(0)&"'); return false;"" style=""border:none;height:21px; width:21px;background: url(imagenes/search.png) no-repeat;"" alt=""Buscar Consulta""></button>")
	
		Response.Write("<button onClick=""cargaDetCamPatExcel('"&Rs(0)&"'); return false;"" style=""border:none;height:21px; width:21px;background: url(imagenes/excel.png) no-repeat;"" alt=""Exportar a Excel""></button></td>")

		for i=k to x
			if (i>=6 and i<=x) then alig="left" else if (i=0) then alig="left" else alig="left" End if End if
		Response.Write("<td  align="&alig&">"&Rs(i)&"</td>")
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")

	rs.close
%>
