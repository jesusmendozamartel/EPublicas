<!--#include file="Conexion.asp"-->
<html xmlns:v="urn:schemas-microsoft-com:vml" 
xmlns:o="urn:schemas-microsoft-com:office:office" 
xmlns:x="urn:schemas-microsoft-com:office:excel" 
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 9">
<!--[if !mso]>
<style>
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
x\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]-->
<!--[if gte mso 9]><xml>
 <o:OfficeDocumentSettings>
  <o:DoNotRelyOnCSS/>
  <o:DoNotUseLongFilenames/>
  <o:DownloadComponents/>
  <o:LocationOfComponents HRef="file:msowc.cab"/>
 </o:OfficeDocumentSettings>
</xml><![endif]-->

<!--[if gte mso 9]><xml>
 <x:ExcelWorkbook>
  <x:ExcelWorksheets>
   <x:ExcelWorksheet>
    <x:Name>BalCom</x:Name>
    <x:WorksheetOptions>
 
     <x:Print>
      <x:ValidPrinterInfo/>
      <x:PaperSizeIndex>9</x:PaperSizeIndex>
      <x:Scale>85</x:Scale>
      <x:HorizontalResolution>600</x:HorizontalResolution>
      <x:VerticalResolution>600</x:VerticalResolution>
     </x:Print>
     <x:Selected/>

    </x:WorksheetOptions>
   </x:ExcelWorksheet>
  </x:ExcelWorksheets>
  <x:WindowHeight>8070</x:WindowHeight>
  <x:WindowWidth>11580</x:WindowWidth>
  <x:WindowTopX>1</x:WindowTopX>
  <x:WindowTopY>1</x:WindowTopY>
  <x:ProtectStructure>False</x:ProtectStructure>
  <x:ProtectWindows>False</x:ProtectWindows>
 </x:ExcelWorkbook>
  <x:ExcelName>
  <x:Name>Print_Titles</x:Name>
  <x:SheetIndex>1</x:SheetIndex>
  <x:Formula>='reporte'!$1:$4</x:Formula>
 </x:ExcelName>
 </xml><![endif]-->
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
dim Tabla1(5000,2500)
Response.Charset= "ISO-8859-1" 
	annio=Request.QueryString("annio")
	nivel=Request.QueryString("nivel")
	codigo=Request.QueryString("codigo")
	'SecSic=Request.QueryString("secsic")
	gcont=Request.QueryString("gcont")
	detalle=Request.QueryString("detalle")

	codText=Request.QueryString("codText")
	detText=Request.QueryString("detText")
	
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

	if detalle =1 then
		codText=codigo
	end if

	Archivo="BalanceComprobacion"
	Titulo="BALANCE DE COMPROBACIÓN, " &annio&" <br>"&NivText&": "&codText&"<br>"&gcont&" / "&detText&"<br> Soles"

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+"_"+annio+".xls"  
	Response.Charset = "UTF-8"
	Response.Write("<table ><tr><td colspan='10' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")

	 response.write("<table width='50%' border='1' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1'  border='1'>")

	if detalle =0 then	
		response.write("<tr><td colspan='6' rowspan='6' align='center' valign='center' style='color:#E3EEF7' bgcolor='#314576'><strong><font size='2pt'>Balance de Comprobación</font></strong></td><td bgcolor='#314576' style='color:#E3EEF7' align='left'>Codigo SICON</td></tr>")
		response.write("<tr bgcolor='#314576'><td  style='color:#E3EEF7' align='left'>Ruc</td></tr>")
		response.write("<tr bgcolor='#314576'><td  style='color:#E3EEF7' align='left'>Razon Social</td></tr>")
		'response.write("<tr bgcolor='#314576'><td style='color:#E3EEF7' align='left'>Ciiu</td></tr>")
		response.write("<tr bgcolor='#314576'><td  style='color:#E3EEF7' align='left'>Ciiu</td></tr>")
		response.write("<tr bgcolor='#314576'><td  style='color:#E3EEF7' align='left'>AE</td></tr>")
		response.write("<tr bgcolor='#314576'><td  style='color:#E3EEF7' align='left'>SECTOR</td></tr>")
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

	Response.Write("<td colspan='3' width='76%'  valign='top'><table class='tabla1' border='1'>")

	SQL="EXEC sp_lista_directorioBC_AnioGContNiv '"&annio&"','"&nivel&"','"&codigo&"','"&gcont&"','"&detalle&"'"
	SQL2=" exec sp_lista_reporteDatosBC_AnioGContNiv '"&annio&"','"&nivel&"','"&codigo&"','"&gcont&"','"&detalle&"'"
		
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con

	X1=cint(rs.fields.count)-1
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

	format ="style='vnd.ms-excel.numberformat:0000;'"

	'if detalle =0 then
	''	X1=X1-1 'TODAS LAS CABECERAS MENOS SECTOR, PORQUE TODOS SON S111
	'end if

	for j=0 to X1

		format =""

		if detalle=0 then
			if j=3 then
				format ="style='vnd.ms-excel.numberformat:0000;'"
			elseif j=4 then
				format ="style='vnd.ms-excel.numberformat:000000;'"
			end if
		else
			if nivel=5 then
				format ="style='vnd.ms-excel.numberformat:000000;'"
			elseif nivel=10 or nivel=11 then
				format ="style='vnd.ms-excel.numberformat:000;'"
			end if
		end if

		response.write("<tr>")
		for i=0 to Y1
			if isnull(Tabla(i,j)) then
				dato="&nbsp;"
			else
				dato=Tabla(i,j)
			end if

			if i Mod 2 = 0 then			
				Response.Write("<td align='center' bgcolor='#ffffff' style='mso-number-format:\@;' >"&dato&"</td>")
			else
				Response.Write("<td align='center' bgcolor='#E3EEF7' style='mso-number-format:\@;'>"&dato&"</td>")
			end if

		next
		 	response.write("</tr>")
	next

	'FRANJA AZUL
	for j=0 to Y1
		
		dato="&nbsp;"
		Response.Write("<td align='center' bgcolor='#314576'>"&dato&"</td>")
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
				''	Response.Write("<td colspan='3' style='color:#E3EEF7' bgcolor='#314576' align='center'><strong>"&dato&"</strong></td>")
				'else
					if IsNumeric(dato) and z>=2143 then
						response.write("<td align='right'>"&FormatNumber(dato,4)&"</td>")
					elseif IsNumeric(dato) then
						Response.Write("<td align='right'>"&FormatNumber(dato,0)&"</td>")
					else
						Response.Write("<td align='right' >"&dato&"</td>")
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

	'Response.Write("<table ><tr><td>&nbsp;&nbsp;</td></tr><tr><tr><td>&nbsp;&nbsp;</td></tr><tr><tr><td colspan='4' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">DETALLE PERSONAL OCUPADO</td></tr></table>")

	'Response.Write("<table class='tabla1'  border='1'><tr bgcolor='#314576' style='color:#E3EEF7'><td colspan='3'>DESCRIPCIÓN</td><td align='left'>CÓDIGO</td></tr>")

	'Response.Write("<td colspan='3' align='left'>AUXILIARES</td><td align='left'>PEO_2</td></tr>")
	'Response.Write("<td colspan='3' align='left'>CONTRATO DE ADMINIST. DE SERVICIOS - CAS</td><td align='left'>PEO_2</td></tr>")
	'Response.Write("<td colspan='3' align='left'>DIRECTORES</td><td align='left'>PEO_1</td></tr>")
	'Response.Write("<td colspan='3' align='left'>DOCENTES UNIVERSITARIOS</td><td align='left'>PEO_2</td></tr>")
	'Response.Write("<td colspan='3' align='left'>EMPLEADO DE CONFIANZA</td><td align='left'>PEO_2</td></tr>")
	'Response.Write("<td colspan='3' align='left'>FUNCIONARIO PÚBLICO</td><td align='left'>PEO_2</td></tr>")
	'Response.Write("<td colspan='3' align='left'>GERENTE</td><td align='left'>PEO_1</td></tr>")
	'Response.Write("<td colspan='3' align='left'>GERENTE GENERAL</td><td align='left'>PEO_1</td></tr>")
	'Response.Write("<td colspan='3' align='left'>GERENTES DE LÍNEA</td><td align='left'>PEO_1</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>JEFES DE AREA/DIVISIÓN</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>OBREROS</td><td align='left'>PEO_3</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>PERSONAL DE LA SALUD</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>PERSONAL DEL MAGISTERIO</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>PERSONAL DIPLOMÁTICO</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>PERSONAL JUDICIAL</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>PERSONAL MILITAR Y POLICIAL</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>PERSONAL PENITENCIARIO</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>PROFESIONALES</td><td align='left'>PEO_8</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>SERVIDOR PÚBLICO - DIRECTOR SUPERIOR</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>SERVIDOR PÚBLICO - EJECUTIVO</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>SERVIDOR PÚBLICO - ESPECIALISTA</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>SERVIDOR PÚBLICO DE APOYO</td><td align='left'>PEO_2</td></tr>")	
	'Response.Write("<td colspan='3' align='left'>TÉCNICOS</td><td align='left'>PEO_3</td></tr>")
	
	'response.write("</tr></table>")
	
	Response.ContentType = "application/save" 

%>
