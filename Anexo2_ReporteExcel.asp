<!--#include file="Conexion.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" 
xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">

<title>Resultado</title>
 <style type="text/css">
<!--
TABLE
{
    BORDER-RIGHT: #c0c0c0 1px dotted;
    BORDER-TOP: #c0c0c0 1px dotted ;
    BORDER-LEFT: #c0c0c0 1px dotted ;
    BORDER-BOTTOM: #c0c0c0 1px dotted;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
}
TABLE.TD
{
    BORDER-RIGHT: #828282 1px dotted ;
    BORDER-TOP: #828282 1px dotted ;
    BORDER-LEFT: #828282 1px dotted ;
    BORDER-BOTTOM: #828282 1px dotted ;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;

}
TD.titulo
{
	BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    background:#E4F2FC;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	HEIGHT: auto;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:right;

}
TD.titulo1
{
	BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	HEIGHT: auto;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:left;


}
TD.act
{
    BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:center;
	HEIGHT:40px;
	width:80px;	
}
TD.dat
{
	BORDER-RIGHT: #828282 1px dotted;
    BORDER-TOP: #828282 1px dotted;
    BORDER-LEFT: #828282 1px dotted;
    BORDER-BOTTOM: #828282 1px dotted;
	PADDING: 0.5em;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:right;
}
TD.cab
{
	BORDER-RIGHT: #828282 1px dotted ;
    BORDER-TOP: #828282 1px dotted ;
    BORDER-LEFT: #828282 1px dotted ;
    BORDER-BOTTOM: #828282 1px dotted ;
	PADDING: 0.5em;
	font-family:Arial, Helvetica, sans-serif;
	font-size:10px;
	VERTICAL-ALIGN: middle;
	text-align:center;
}
-->
</style>
</head>
<body >
<%
dim Tabla(5000,5000)
dim Tabla1(5000,5000)

Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	nivel=Request.QueryString("nivel")
	codigo=Request.QueryString("codigo")
	gcont=Request.QueryString("gcont")
	detalle=Request.QueryString("detalle")
	codText=Request.QueryString("codText")
	detText=Request.QueryString("detText")
	eeff=Request.QueryString("eeff")
	moneda=Request.QueryString("moneda")

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

MonedaText=	""
if moneda ="0" then
	MonedaText="SOLES"
elseif moneda ="1" then
	MonedaText="MILES DE SOLES"
elseif moneda ="2" then
	MonedaText="MILLONES DE SOLES"
end if	
'-----------------------------	
Titulo1="Anexo 2 - Existencias / Bienes Realizables"
Titulo2=Titulo1&" <br>" &annio&" <br>"&NivText&": "&codText&"<br>"&gcont&" / "&detText&"<br>"&MonedaText&"<br>"
tablaname="anexo2_res"
campos=114
ntabla=12
'------------------------------	

	SQL="exec sp_lista_ReglasBC_Anexo "&annio&","&ntabla&" "
	
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	Response.Buffer = true
	response.AddHeader "Content-Disposition", "attachment; filename="+eeff+"_"+annio+".xls" 
	Response.Write("<table ><tr><td colspan='10' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo2&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")
	response.write("<table bgcolor='#FFFFFF' width='100%' border='0' cellspacing='0' cellpadding='0'><tr><td width='50%' valign='top'><table  class='tabla1' width='100%' border='0'>")
	if detalle =0 then
		response.write("<tr bgcolor='#314576' style='color:#E3EEF7'><td colspan='2' rowspan='6' align='center' style='vertical-align:middle;'><strong><font size='2pt'  style='color:#E3EEF7'>"&Titulo1&"</font></strong></td><td>Cod_Empresa</td></tr>")
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
		if rs(0) ="001" or rs(0) ="013" or rs(0) ="025" or rs(0) ="037" or rs(0) ="049" or rs(0) ="061"  or rs(0) ="073" or rs(0) ="085" or rs(0) ="097" or rs(0) ="110"  then
				response.write("<tr bgcolor='#FFE296' style=""font-weight:bold"">")
		elseif rs(0) ="111" or rs(0) ="113" then 'PARA PINTARLO DE NEGRITA
				response.write("<tr style=""font-weight:bold"">")
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
		        '---------- AQUI PINTA EL REGISTRO DE OTRO COLOR --------------------
				if j =0 or j =12 or j =24 or j =36 or j =48 or j =60 or j =72 or j =84 or j =96 or j =109 then
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
	
	Response.ContentType = "application/save" 
%>
