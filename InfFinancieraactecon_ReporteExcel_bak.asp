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
    <x:Name>EstGanPer</x:Name>
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
	width:40px;
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
	width:40px;
	HEIGHT:40px;
}
-->
</style>
</head>
<body >
<%
dim Tabla(5000,5000)
dim Tabla1(5000,5000)
Response.Charset= "ISO-8859-1" 
grupo=Request.QueryString("grupo") 
annio=Request.QueryString("annio")
nivel=Request.QueryString("nivel")

IF grupo = "G1 G3" THEN grupo="G1+G3" END IF
IF grupo = "G2 G4" THEN grupo="G2+G4" END IF


SQL="select Clave_cta,Desc_cta, campo from Sunatma_glosas order by Clave_cta"
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
   	rs.Open SQL, con

	Archivo="InfFinancieraActEcon_" &annio
	Titulo="SUNAT - Información Financiera por Actividad Económica" &annio

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+".xls" 
	Response.Charset= "ISO-8859-1" 	
	Response.Write("<table><tr><td colspan='6' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")			
	Response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0' ><tr><td width='24%' valign='top'><table  border='1'>")
	Response.write("<tr bgcolor='#FFE2C6'><td colspan='2' rowspan='2' valign='center' class='act' ><strong><font size='2pt'>Información Financiera</font></strong></td><td align='right' class='act' ><strong>Cod_ActEcon</strong></td></tr>")
	 response.write("<tr bgcolor='#FFE2C6'><td class='act'><strong>Descripción</strong></td></tr>")
	 response.write("<tr bgcolor='#FFE2C6'><td class='act'><strong>Clave</strong></td><td class='act'><strong>Descripción</strong></td><td align='center' class='act'><strong>Campo</strong></td></tr>")
    	while not rs.eof
		if rs(0)="001" or rs(0)="002" or rs(0)="035" or rs(0)="048" or rs(0)="061"then bgcolor="#DADBDB" else bgcolor="" end if 
	    response.write("<tr bgcolor='"&bgcolor&"'><td STYLE='vnd.ms-excel.numberformat:@' class='titulo1'>"&rs(0)&"</td><td STYLE='vnd.ms-excel.numberformat:@' class='titulo1'>"&rs(1)&"</td><td STYLE='vnd.ms-excel.numberformat:@' class='titulo1'>"&rs(2)&"</td></tr>")

		rs.MoveNext
		wend
		rs.Close
		Set rs=Nothing
	response.write("</table></td>")
	
	response.write("<td width='76%' valign='top' ><table  border='1'>")

	SQL="sp_listadatosACTECON '"&annio& "','"&grupo& "','"&nivel&"','1'" 'cabecera
	SQL2="sp_listadatosACTECON '"&annio& "','"&grupo& "','"&nivel&"','2'"'cuerpo
	
''	response.Write(SQL2)
''	response.End()


	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
    rs.Open sql, con
	X1=cint(RS.fields.count)-1
	Y1=cint(rs.RecordCount )-1
	Session.LCID = 1034
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
					response.write("<td colspan='1' class='cab' bgcolor='#FFE2C6' STYLE='vnd.ms-excel.numberformat:@'>"&Tabla(i,j)&"</td>")
			else
					response.write("<td colspan='1' class='cab' bgcolor='#FFE2C6' STYLE='vnd.ms-excel.numberformat:@'>"&Tabla(i,j)&"</td>")
			end if
		next
		 	response.write("</tr>")
	next

	Set rs2 = Server.CreateObject("ADODB.Recordset")	
	rs2.CursorLocation=3
    rs2.Open sql2, con
	X2=cint(RS2.fields.count)-1
	Y2=cint(rs2.RecordCount )-1
	'response.write(X2&"-"&Y2)
	i=0
	Session.LCID = 1034
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
					response.write("<td bgcolor='#FFE2C6' align='center' class='dat'><strong>"&Tabla1(i,j)&"</strong></td>")
				else
					if j = 1 or j = 2 or j = 35 or j=48 or j=61 then 
					response.write("<td bgcolor='#DADBDB' align='right' class='dat'>&nbsp;</td>")
					else
					response.write("<td align='right' class='dat'>"&Tabla1(i,j)&"</td>")
					end if
				end if

		next
		response.write("</tr>")
	next
	response.write("</table></td></tr></table>")
	Response.ContentType = "application/save" 

%>
