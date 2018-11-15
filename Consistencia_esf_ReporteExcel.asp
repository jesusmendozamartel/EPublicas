<!--#include file="Conexion.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" 
xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">
<style type="text/css">
body {
	margin-left: 0px;
	margin-right: 0px;
	margin-top: 0px;
	margin-bottom: 0px;
	background-image: url(Imagenes/fdopag.jpg);
}
TABLE.tabla1
{
    BORDER-RIGHT: #314576 1px solid;
    BORDER-TOP: #314576 1px solid;
    BORDER-LEFT: #314576 1px solid;
    BORDER-BOTTOM: #314576 1px solid;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
	background:#FFFFFF;
}

TABLE.tabla1 TD
{
    BORDER-RIGHT: #314576 1px solid;
    BORDER-TOP: #314576 1px solid;
    BORDER-LEFT: #314576 1px solid;
    BORDER-BOTTOM: #314576 1px solid;
	
}
TABLE.tabla1 TH
{
    BORDER-RIGHT: #314576 1px solid;
    PADDING-RIGHT: 5px;
    BORDER-TOP: #314576 1px solid;
    PADDING-LEFT: 5px;
	background:#314576;
    PADDING-BOTTOM: 5px;
    BORDER-LEFT: #314576 1px solid;
    PADDING-TOP: 5px;
    BORDER-BOTTOM: #314576 1px solid;
    HEIGHT: 20px;
	color:#E3EEF7;
	font-family:Arial, Helvetica, sans-serif;
	font-size:12px;
	
}
</style>
<title>Resultado</title></head>
<body >
<%
	Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")

	Archivo="Consistencia_esf"
	Titulo="CONSISTENCIAS - CUENTAS DEL BALANCE DE COMPROBACIÓN (ESTADO DE SITUACIÓN FINANCIERA) Y ESTADO FINANCIERO "&annio
	
	SQL="sp_lista_consistencias '"&annio& "' " 'Lista los registros solicitados
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con 

	x=rs.Fields.Count-1

	if rs.RecordCount=1 then
		Response.Write(rs.RecordCount) 'No se encontraron registros!
		Response.End
	End if
	
	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+"_"+annio+".xls" 
	Response.Charset= "ISO-8859-1" 	

	Response.Write("<table ><tr><td colspan='26' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")
	Response.Write("<br>")
	Response.Write("<table>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'><br>Cantidad de Empresas Públicas: "&rs.RecordCount&"</td>")
	Response.Write("</tr>")
	Response.Write("</table>")
	Response.Write("<br>")

	response.write("<table bgcolor='#FFFFFF' width='100%' border='0' cellspacing='0' cellpadding='0' class='tabla1'> ")
	response.write("<tr bgcolor='#787a82' style='color:#E3EEF7'>")
	response.write("<td width='8%' rowspan='3'><div align='center'>RUC</div></td>")
	response.write("<td width='25%' rowspan='3'>RAZ&Oacute;N SOCIAL</td>")
	response.write("<td width='5%' rowspan='3'><div align='center'>C&Oacute;DIGO <br />SICON</div></td>")
	response.write("<td width='5%' rowspan='3'><div align='center'>SECTOR <br />SICON</div></td>")
	response.write("<td width='5%' rowspan='3'><div align='center'>CIIU</div></td>")
	response.write("<td width='5%' rowspan='3'><div align='center'>AE</div></td>")
	response.write("<td bgcolor='#76819a' colspan='10'><div align='center'>BALANCE DE COMPROBACI&Oacute;N</div></td>")
	response.write("<td bgcolor='#314576' colspan='10'><div align='center'>ESTADO FINANCIERO (ESTADO DE SITUACI&Oacute;N FINANCIERA)</div></td>")
	response.write("</tr>")
	response.write("<tr bgcolor='#314576' style='color:#E3EEF7'>")
	response.write("<td colspan=5' bgcolor='#76819a'><div align='center'>A&Ntilde;O "&annio&"</div></td>")
	response.write("<td colspan='5' bgcolor='#76819a'><div align='center'>A&Ntilde;O "&annio-1&"</div></td>")
	response.write("<td colspan='5'><div align='center'>A&Ntilde;O "&annio&"</div></td>")
	response.write("<td colspan='5'><div align='center'>A&Ntilde;O "&annio-1&"</div></td>")
	response.write("</tr>")
	response.write("<tr bgcolor='#314576' style='color:#E3EEF7'>")
	response.write("<td width='2%' height='89' bgcolor='#76819a'><div align='center'>TOTAL<br />ACTIVO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>TOTAL<br />PASIVO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>TOTAL<br />PATRIMONIO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>TOTAL PASIVO+ <br />TOTAL PATRIMONIO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>DIFERENCIA</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'height='89'><div align='center'>TOTAL<br />ACTIVO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>TOTAL<br />PASIVO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>TOTAL<br />PATRIMONIO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>TOTAL PASIVO+ <br />TOTAL PATRIMONIO</div></td>")
	response.write("<td width='2%' bgcolor='#76819a'><div align='center'>DIFERENCIA</div></td>")
	response.write("<td width='2%' height='89'><div align='center'>TOTAL<br />ACTIVO</div></td>")    
	response.write("<td width='2%'><div align='center'>TOTAL<br />PASIVO</div></td>")
	response.write("<td width='2%'><div align='center'>TOTAL<br />PATRIMONIO</div></td>")
	response.write("<td width='2%'><div align='center'>TOTAL PASIVO+ <br />TOTAL PATRIMONIO</div></td>")
	response.write("<td width='2%'><div align='center'>DIFERENCIA</div></td>")        
	response.write("<td width='2%' height='89'><div align='center'>TOTAL<br />ACTIVO</div></td>")
	response.write("<td width='2%'><div align='center'>TOTAL<br />PASIVO</div></td>")
	response.write("<td width='2%'><div align='center'>TOTAL<br />PATRIMONIO</div></td>")
	response.write("<td width='2%'><div align='center'>TOTAL PASIVO+ <br />TOTAL PATRIMONIO</div></td>")
	response.write("<td width='2%'><div align='center'>DIFERENCIA</div></td>")	
	response.write("</tr>")

	'-----------------------------------------------------------------------------------------------------------------------------------------------------
	j=0
	while not rs.eof
		if j=0 then bg="bgcolor='#FFFFFF'" else bg="" End if
		Response.Write("<tr "&bg&">")
	
		for i=k to x
			if (i>=6 and i<=x) then alig="left" else if (i=0) then alig="left" else alig="left" End if End if
		Response.Write("<td STYLE='vnd.ms-excel.numberformat:@'  align="&alig&">"&Rs(i)&"</td>")
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")

	rs.close

%>
</body >
</html>
