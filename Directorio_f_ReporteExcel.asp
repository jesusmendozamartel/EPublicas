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
    BORDER-RIGHT: #314576 1px dotted;
    BORDER-TOP: #314576 1px dotted;
    BORDER-LEFT: #314576 1px dotted;
    BORDER-BOTTOM: #314576 1px dotted;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
	background:#FFFFFF;
}

TABLE.tabla1 TD
{
    BORDER-RIGHT: #314576 1px dotted;
    BORDER-TOP: #314576 1px dotted;
    BORDER-LEFT: #314576 1px dotted;
    BORDER-BOTTOM: #314576 1px dotted;
	
}
TABLE.tabla1 TH
{
    BORDER-RIGHT: #314576 1px dotted;
    PADDING-RIGHT: 5px;
    BORDER-TOP: #314576 1px dotted;
    PADDING-LEFT: 5px;
	background:#314576;
    PADDING-BOTTOM: 5px;
    BORDER-LEFT: #314576 1px dotted;
    PADDING-TOP: 5px;
    BORDER-BOTTOM: #314576 1px dotted;
    HEIGHT: 20px;
	color:#E3EEF7;
	font-family:Arial, Helvetica, sans-serif;
	font-size:12px;
	
}
</style>
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
    <x:Name>Directorio Fonafe</x:Name>
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
 
<title>Resultado</title></head>
<body >
<%
	Response.Charset= "ISO-8859-1" 

	Archivo="Directorio_fonafe"
	Titulo="Directorio FONAFE"
	
	SQL="sp_lista_directorio_Anio_FONAFE"
	Set rs = Server.CreateObject("ADODB.Recordset")	
	rs.CursorLocation=3
	rs.Open SQL, con 

	x=rs.Fields.Count-1
	
	if rs.RecordCount=1 then
		Response.Write(rs.RecordCount) ''No se encontraron registros!
		Response.End
	End if
	j=0

    Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+".xls" 
	Response.Charset= "ISO-8859-1" 	
	
	Response.Write("<table ><tr><td colspan='16' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")

	Response.Write("<br>")
	Response.Write("<table>")
	Response.Write("<tr bgcolor='#FFFFFF'>")
	Response.Write("<td colspan='4' align='left'><br><strong>Cantidad de Empresas: "&rs.RecordCount&"</strong></td>")
	Response.Write("</tr>")
	Response.Write("</table>")
	Response.Write("<br>")

	Response.Write("</td></tr>")

	response.write("<table  class='tabla1'  border='0'>")

	for i=0 to x 
		Response.Write("<th bgcolor='#314576' >"&rs.fields(i).name&"</th>")
	next

	while not rs.eof
		Response.Write("<tr>")
	
		for i=k to x
			if (i>=6 and i<=x) then alig="center" else if (i=0) then alig="center" else alig="left" End if End if
		Response.Write("<td STYLE='vnd.ms-excel.numberformat:@' align="&alig&">"&Rs(i)&"</td>")
	
		next
		Response.Write("</tr>")
		rs.MoveNext
		j=j+1
	wend
	Response.Write("</table>")
	response.write("</tr></table>")
	Response.ContentType = "application/save" 

%>
</body >
</html>
