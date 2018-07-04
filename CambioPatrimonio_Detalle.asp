<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,2000)
Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	CodiEnt=Request.QueryString("CodiEnt")

	SQL="EXEC sp_lista_DetalleCamPat_AnioCodiEnt "&annio&","&CodiEnt

	Set rs = Server.CreateObject("ADODB.Recordset")	

	rs.CursorLocation=3
	rs.Open SQL, con 
	
	x=rs.Fields.Count-1
	
	j=0

	Response.Write("<br>")
	Response.Write("<table class='tabla1'>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td align='left'>Codigo Entidad: "&Rs(0)&"</td>")
	Response.Write("<td align='left'>RUC: "&Rs(1)&"</td>")
	Response.Write("<td align='left'>ENTIDAD: "&Rs(2)&"</td>")
	Response.Write("<td align='left'>Ciiu_R4_4d: "&Rs(3)&"</td>")
	Response.Write("<td align='left'>AE: "&Rs(4)&"</td>")
	Response.Write("<td align='left'>SECTOR: "&Rs(5)&"</td>")
	Response.Write("</tr>")
	Response.Write("</table>")

	Response.Write("<table class='tabla1'>")

	for i=6 to x 
		Response.Write("<th >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		for i=6 to x
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
