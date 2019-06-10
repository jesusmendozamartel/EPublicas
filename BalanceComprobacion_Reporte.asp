<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,2500)
Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	nivel=Request.QueryString("nivel")
	codigo=Request.QueryString("codigo")
	'SecSic=Request.QueryString("secsic")
	gcont=Request.QueryString("gcont")
	detalle=Request.QueryString("detalle")

	'SQL=" exec sp_lista_cuentas_RepAnioSecSicon '01',"&annio&",'"&SecSic&"'"

	SQL=" exec sp_lista_ReglasBC_Anio "&annio

	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
   	rs.Open SQL, con
	
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


	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1'  border='1'>")

	if detalle =0 then		
		response.write("<tr><td colspan='6' rowspan='6' align='center'  style='color:#E3EEF7' bgcolor='#314576'><strong><font size='2pt'>Balance de Comprobación</font></strong></td><td colspan='3' bgcolor='#314576' style='color:#E3EEF7' align='left'>Codigo SICON</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='left'>Ruc</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='left'>Razon Social</td></tr>")
		'response.write("<tr bgcolor='#314576'><td colspan='3' style='color:#E3EEF7' align='left'>Ciiu</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='left'>Ciiu</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='left'>AE</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='left'>SECTOR</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' >		NroOrden</td><td  style='color:#E3EEF7' align='center'>FActivo</td><td align='left' style='color:#E3EEF7' >FPasivo</td><td align='left' style='color:#E3EEF7' >SaldoActivo</td><td align='left' style='color:#E3EEF7' >SaldoPasivo</td><td align='left' style='color:#E3EEF7' >CodCta</td><td align='left' style='color:#E3EEF7' >Descripcion</td></tr>")

	elseif detalle =1 then
		response.write("<tr><td colspan='4' rowspan='2' align='center' style='color:#E3EEF7' bgcolor='#314576'><strong><font size='1pt'>Balance de Comprobación</font></strong></td><td colspan='3' style='color:#E3EEF7' bgcolor='#314576' align='right'>&nbsp;</td></tr>")
		response.write("<tr><td colspan='3' style='color:#E3EEF7' bgcolor='#314576' align='right'>"&NivText&"</td></tr>")
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td>NroOrden</td><td align='center'>FActivo</td><td align='left'>FPasivo</td><td align='left'>SaldoActivo</td><td align='left'>SaldoPasivo</td><td align='left'>CodCta</td><td align='left'>Descripcion</td></tr>")
	end if

	while not rs.eof
		response.write("<tr><td align='center'>"&rs(0)&"</td><td>"&rs(1)&"</td><td align='left'>"&rs(2)&"</td><td align='left'>"&rs(3)&"</td><td align='left'>"&rs(4)&"</td><td align='left'>"&rs(5)&"</td><td align='left'>"&rs(6)&"</td></tr>")
    	rs.MoveNext
	wend
	rs.Close
	Set rs=Nothing
	response.write("</table></td>")

	response.write("<td width='76%'  valign='top'><table class='tabla1' border='0'>")

	SQL="EXEC sp_lista_directorioBC_AnioGContNiv '"&annio&"','"&nivel&"','"&codigo&"','"&gcont&"','"&detalle&"'"
	SQL2=" exec sp_lista_reporteDatosBC_AnioGContNiv '"&annio&"','"&nivel&"','"&codigo&"','"&gcont&"','"&detalle&"'"

	'response.write(SQL)
	'response.write(SQL2)
	'response.end	

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

	'if detalle =0 then
	''	X1=X1-1 'TODAS LAS CABECERAS MENOS SECTOR, PORQUE TODOS SON S111
	'end if

	for j=0 to X1
		response.write("<tr>")
		for i=0 to Y1
			if isnull(Tabla(i,j)) then
				dato="&nbsp;"
			else
				dato=Tabla(i,j)
			end if

			if i Mod 2 = 0 then
				response.write("<td align='center' bgcolor='#ffffff' >"&dato&"</td>")
				
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

	Set rs2 = Server.CreateObject("ADODB.Recordset")	
	rs2.CursorLocation=3
    rs2.Open sql2, con

    z=1

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
		response.write("<tr>")
			for i=0 to Y2
				if isnull(Tabla1(i,j)) then
					dato="&nbsp;"
				else
					dato=Tabla1(i,j)
				end if

				'if j=0 then
				''	response.write("<td style='color:#E3EEF7' bgcolor='#314576' align='center'><strong>"&dato&"</strong></td>")
				'else
					if IsNumeric(dato) and z>=2143 then
						response.write("<td align='right'>"&FormatNumber(dato,4)&"</td>")
					elseif IsNumeric(dato) then
						response.write("<td align='right'>"&FormatNumber(dato,0)&"</td>")
					else
						response.write("<td align='right' >"&dato&"</td>")
					end if
					
				'end if
			next
			'Response.Flush
			response.write("</tr>")
			z=z+1
		next

		Set rs2 = rs2.NextRecordset
	wend

	response.write("</table></td></tr></table>")
	
%>
