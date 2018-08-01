<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,5000)

Response.Charset= "ISO-8859-1" 
annio=Request.QueryString("annio")
nivel=Request.QueryString("nivel")
codigo=Request.QueryString("codigo")
gcont=Request.QueryString("gcont")
detalle=Request.QueryString("detalle")
moneda=Request.QueryString("moneda")
'----------------
Titulo1="ANEXO 6 - DEPRECIACIÓN ACUMULADA DE INMUEBLES, MAQUINARIA Y EQUIPO"
tablaname="anexo6_res"
campos=110
ntabla=14
'---------------------
	SQL="exec sp_lista_ReglasBC_Anexo "&annio&","&ntabla&" "
	
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con

response.write("<table bgcolor='#FFFFFF' width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td width='50%' valign='top'><table  class='tabla1' width='100%' border='0'>")

	NivText=	""
	if nivel =5 then
		NivText="ACTIVIDAD"
	elseif nivel =6 then
		NivText="SECTOR INSTITUCIONAL"
	elseif nivel =10 then
		NivText="ACTIVIDAD 54"
	elseif nivel =11 then
		NivText="ACTIVIDAD 14"
	elseif nivel =12 then
		NivText="SECTOR SICON"
	end if
	
	if detalle =0 then
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td colspan='2' rowspan='6' align='center'><strong><font size='2pt'  style='color:#E3EEF7'>"&Titulo1&"</font></strong></td><td>Cod_Empresa</td></tr>")
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td>RUC</td></tr>")
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td>Razón Social</td></tr>")
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td>Ciiu</td></tr>")
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td>Actividad</td></tr>")
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td>Sector</td></tr>")
	elseif detalle =1 then
		response.write("<tr><td colspan='2' rowspan='2' align='center' style='color:#E3EEF7' bgcolor='#314576'><strong><font size='1pt'>"&Titulo1&"</font></strong></td><td  style='color:#E3EEF7' bgcolor='#314576' align='right'>&nbsp;</td></tr>")
		response.write("<tr><td style='color:#E3EEF7' bgcolor='#314576' align='right'>"&NivText&"</td></tr>")
	end if
	
	response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td align='center'>Nro Orden</td><td align='center'>Clave</td><td align='left'>Descripción</td></tr>")
	'BG X
	while not rs.eof	' PARA PINTARLO DE NEGRITA Y DE FONDO NARANJA
		if rs(0) ="001" or rs(0) ="012" or rs(0) ="023" or rs(0) ="034" or rs(0) ="045" or rs(0) ="056"  or rs(0) ="067" or rs(0) ="078" or rs(0) ="089" or rs(0) ="100"  then
				response.write("<tr bgcolor='#FFE296' style=""font-weight:bold"">")
		else 
				response.write("<tr>")	
		end if

		response.write("<td align='center'>"&rs(0)&"</td><td align='center'>"&rs(1)&"</td><td align='left'>"&rs(2)&"</td></tr>")
		rs.MoveNext	
	wend
	
	rs.Close
	Set rs=Nothing

	response.write("</table></td>")
	response.write("<td width='50%' valign='top'><table class='tabla1' border='0'>")

   	SQL= "EXEC sp_lista_anexos "&campos&",'"&annio&"','"&nivel& "','"&codigo& "','"&gcont&"','1',"&tablaname&","&detalle&","&moneda&"" 'cabecera
	
	SQL2="EXEC sp_lista_anexos "&campos&",'"&annio& "','"&nivel& "','"&codigo& "','"&gcont&"','2',"&tablaname&","&detalle&","&moneda&"" 'detalle
'-------------------------------------------------------------------------------------------------------------------------------------------------	
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
			if isnull(Tabla(i,j)) then
				dato="&nbsp;"
			else
				dato=Tabla(i,j)
			end if

			if i Mod 2 = 0 then
				response.write("<td align='center' bgcolor='#314576' style='color:#E3EEF7'>"&dato&"</td>")
			else
				response.write("<td align='center' bgcolor='#E3EEF7'>"&dato&"</td>")
			end if
		next
		 	response.write("</tr>")
	next

	'FRANJA AZUL
	for j=0 to Y1
		
		dato="&nbsp;"
		response.write("<td align='center' bgcolor='#314576'>"&dato&"</td>")
	next
'-------------------------------------------------------------------------------------------------------------------------------------------------	
	'BG X
	Set rs2 = Server.CreateObject("ADODB.Recordset")	
		rs2.CursorLocation=3
		rs2.Open sql2, con

	while not rs2 is Nothing

		Erase Tabla1' clear the array

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

		for j=0 to X2
		        '---------- AQUI PINTA EL REGISTRO DE OTRO COLOR --------------------
				if j =0 or j =11 or j =22 or j =33 or j =44 or j =55 or j =66 or j =77 or j =88 or j =99 then
					response.write("<tr bgcolor='#FFE296' style=""font-weight:bold"">")
				else
					response.write("<tr>")
				end if
				'--------------------------------------------------------------------
			for i=0 to Y2
				if isnull(Tabla1(i,j)) then
					dato="&nbsp;"
				else
					dato=Tabla1(i,j)
				end if

				if IsNumeric(dato) and moneda="0" and dato>"0" then
					response.write("<td align='right'>"&FormatNumber(dato,0)&"</td>")
				elseif IsNumeric(dato) and moneda="1" and dato>"0" then
					response.write("<td align='right'>"&FormatNumber(dato,3)&"</td>")
				elseif IsNumeric(dato) and moneda="2" and dato>"0" then
					response.write("<td align='right'>"&FormatNumber(dato,6)&"</td>")
				else
					Response.Write("<td align='right'>"&dato&"</td>")
				end if
			next
			'Response.Flush
			response.write("</tr>")
			
		next

		Set rs2 = rs2.NextRecordset
	wend

	response.write("</table></td></tr></table>")
%>
