<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,2000)
Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	nivel=Request.QueryString("nivel")
	codigo=Request.QueryString("codigo")
	detalle=Request.QueryString("detalle")
	xtrim=Request.QueryString("xtrim")


	SQL=" exec [sp_lista_cuentas_fonafe] '23'"

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
	end if

	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1'  border='1'>")

	if detalle =0 then		
		response.write("<tr><td colspan='2' rowspan='5' align='center'  style='color:#E3EEF7' bgcolor='#314576'><strong><font size='2pt'>Estado de resultados integrales</font></strong></td><td bgcolor='#314576' style='color:#E3EEF7' align='right'>Codigo</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='right'>Ruc</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='right'>Razon Social</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='right'>Ciiu</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='right'>AE</td></tr>")
		response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' colspan='2' >NroOrden</td><td align='left' style='color:#E3EEF7' >Descripcion</td></tr>")
	elseif detalle =1 then
		response.write("<tr><td colspan='2' rowspan='2' align='center' style='color:#E3EEF7' bgcolor='#314576'><strong><font size='1pt'>Estado de resultados integrales</font></strong></td><td  style='color:#E3EEF7' bgcolor='#314576' align='right'>&nbsp;</td></tr>")
		response.write("<tr><td style='color:#E3EEF7' bgcolor='#314576' align='right'>"&NivText&"</td></tr>")
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td colspan='2'>NroOrden</td><td align='left'>Descripcion</td></tr>")
	end if

	while not rs.eof
		response.write("<tr><td align='center' colspan='2'>"&rs(0)&"</td><td>"&rs(1)&"</td></tr>")
    rs.MoveNext
	wend
	rs.Close
	Set rs=Nothing
	response.write("</table></td>")

	response.write("<td width='76%'  valign='top'><table class='tabla1' border='0'>")

	SQL="EXEC sp_lista_directorio_fonafe '"&nivel&"','ERI_PRO','"&annio&"','"&xtrim&"','"&codigo&"','"&detalle&"'"
	SQL2=" exec sp_lista_reporteDatos_fonafe '37','ERI_PRO','"&annio&"','"&xtrim&"','"&nivel&"','"&codigo&"','"&detalle&"'"

'response.Write(sql)
'response.Write(sql2)
'response.End()

Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
    rs.Open sql, con

Set rs2 = Server.CreateObject("ADODB.Recordset")	
	rs2.CursorLocation=3
    rs2.Open sql2, con		

'if (rs2.RecordCount=0) then
'	response.Write("<strong>¡No se encontraron datos!</strong>")
'	response.End()
'end if

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
				response.write("<td colspan='3' align='center' bgcolor='#ffffff'>"&dato&"</td>")
				
			else
				response.write("<td colspan='3' align='center' bgcolor='#E3EEF7'>"&dato&"</td>")
				
			end if
		next
		 	response.write("</tr>")
	next
'


	X2=cint(RS2.fields.count)-1
	Y2=cint(rs2.RecordCount )-1
	'response.write(X2&"-hhh"&Y2)
'    'response.End()
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
'	
	for j=0 to X2
	response.write("<tr>")
		for i=0 to Y2
					if isnull(Tabla1(i,j)) then
						dato="&nbsp;"
					else
						dato=Tabla1(i,j)
					end if

			if j=0 then
				response.write("<td style='color:#E3EEF7' bgcolor='#314576' align='center'><strong>"&dato&"</strong></td>")
			else
				if IsNumeric(dato) then
					response.write("<td align='right'>"&FormatNumber(dato,0)&"</td>")
				else
					response.write("<td align='right' >"&dato&"</td>")
				end if
				
			end if
		next
		'Response.Flush
		response.write("</tr>")
	next
'	
	response.write("</table></td></tr></table>")

	
%>
