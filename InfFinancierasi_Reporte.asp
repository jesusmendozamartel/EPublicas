<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,5000)

Response.Charset= "ISO-8859-1" 
grupo=Request.QueryString("grupo") 
annio=Request.QueryString("annio")
IF grupo = "G1 G3" THEN grupo="G1+G3" END IF
IF grupo = "G2 G4" THEN grupo="G2+G4" END IF



SQL="select Clave_cta,Desc_cta, campo from Sunatma_glosas order by Clave_cta"
Set rs = Server.CreateObject("ADODB.Recordset")	
rs.CursorLocation=3
rs.Open SQL, con


response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")
	response.write("<tr bgcolor='#FFE2C6'><td colspan='2' rowspan='2' align='center'><strong><font size='2pt'>Información Financiera</font></strong></td><td><strong>Cod_Sector</strong></td></tr>")
	response.write("<tr bgcolor='#FFE2C6'><td><strong>Descripción</strong></td></tr>")
	response.write("<tr bgcolor='#FFE2C6'><td><strong>Clave</strong></td><td><strong>Descripción</strong></td><td><strong>Campo</strong></td></tr>")
	while not rs.eof
	if rs(0)="001" or rs(0)="002" or rs(0)="035" or rs(0)="048" or rs(0)="061"then bgcolor="#DADBDB" else bgcolor="" end if 
	response.write("<tr bgcolor='"&bgcolor&"'><td>"&rs(0)&"</td><td>"&rs(1)&"</td><td>"&rs(2)&"</td></tr>")
	rs.MoveNext
	wend
	rs.Close
	Set rs=Nothing
	response.write("</table></td>")
	response.write("<td width='76%' valign='top'><table class='tabla1' border='0'>")


	SQL="sp_listadatosSI '"&annio& "','"&grupo& "','1'" 'cabecera
	SQL2="sp_listadatosSI '"&annio& "','"&grupo& "','2'"'cuerpo
		
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open sql, con
	X1=cint(RS.fields.count)-1
	Y1=cint(rs.RecordCount )-1
	while not rs.eof
	   for j=0 to X1
		 Tabla(i,j)=rs(j)
	   next
	rs.MoveNext
	i=i+1
	wend 
	rs.Close
	Set rs=Nothing
	for j=0 to X1
	response.write("<tr>")
		for i=0 to Y1
			if i Mod 2 = 0 then
			response.write("<td colspan='1' align='center' bgcolor='#FFE2C6'>"&Tabla(i,j)&"</td>")
			else
			response.write("<td colspan='1' align='center' bgcolor='#FFE2C6'>"&Tabla(i,j)&"</td>")
			end if
		next
			response.write("</tr>")
	next
	Set rs2 = Server.CreateObject("ADODB.Recordset")	
		rs2.CursorLocation=3
		rs2.Open sql2, con
		X2=cint(RS2.fields.count)-1
		Y2=cint(rs2.RecordCount )-1
		i=0
		 while not rs2.eof
		   for j=0 to X2
			 Tabla1(i,j)=rs2(j)
			next
   		    rs2.MoveNext
			i=i+1
		 wend 
		 rs2.Close
		 Set rs2=Nothing
				
		 for j=0 to X2
		 response.write("<tr>")
			for i=0 to Y2
				if j=0 then 
					response.write("<td bgcolor='#FFE2C6' align='center'><strong>"&Tabla1(i,j)&"</strong></td>")
				else
					if j = 1 or j = 2 or j = 35 or j=48 or j=61 then 
					response.write("<td bgcolor='#DADBDB' align='right'>&nbsp;</td>")
					else
					response.write("<td align='right' >"&Tabla1(i,j)&"</td>")
					end if
				end if
			next
			response.write("</tr>")
		next
		response.write("</table></td></tr></table>")

	RESPONSE.Write("  </table>")
	RESPONSE.Write("  </table>")
	


%>
