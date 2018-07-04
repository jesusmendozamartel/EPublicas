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
    <x:Name>EstFin</x:Name>
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

	
	anio=Request.QueryString("anio")
	formato=Request.QueryString("for")
	grupo=Request.QueryString("strGru")
	sector=Request.QueryString("sec")
	ForText=Request.QueryString("ForText")	
	SosText=Request.QueryString("SosText")	
	GruText=Request.QueryString("GruText")	

	Dim rsC
	Set rsC=Server.CreateObject("ADODB.RecordSet")	
	rsC.CursorLocation=3

	rsC.Open "exec sp_EstadosFinancieros_ListarPivot_AnioFormatoSectorGContable2_Columns "&anio&",'"&formato&"','"&sector&"','"&grupo&"'", Con


	Dim rsR
	Set rsR=Server.CreateObject("ADODB.RecordSet")	
	rsR.CursorLocation=3

	rsR.Open "exec sp_EstadosFinancieros_ListarPivot_AnioFormatoSectorGContable2_Rows "&anio&",'"&formato&"','"&sector&"','"&grupo&"'", Con
	

	Dim rsD
	Set rsD=Server.CreateObject("ADODB.RecordSet")	
	rsD.CursorLocation=3

	rsD.Open "exec sp_EstadosFinancieros_ListarPivot_AnioFormatoSectorGContable2_Data "&anio&",'"&formato&"','"&sector&"','"&grupo&"'", Con

	Archivo="ConsEstadosFinancierosSICON" &annio
	Titulo="SICON - Estados Financieros "&anio&"/"&ForText&"/"&SosText&"/"&sector&"/"&GruText

	Response.Charset = "UTF-8"
	response.ContentType = "application/vnd.ms-excel" 
	response.AddHeader "Content-Disposition", "attachment; filename="+Archivo+".xls" 
	Response.Charset= "ISO-8859-1" 	
	Response.Write("<table><tr><td colspan='6' align='center'  style=""font-family:Arial, Helvetica, sans-serif; font-size:20px; color:#003300"">"&Titulo&"</td></tr><tr><td>&nbsp;&nbsp;</td></tr><tr>")			

	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")


Set objFieldsC = rsC.Fields
	response.write("<tr bgcolor='#FFE2C6'><td rowspan=4><strong>"&ForText&"</strong></td>")

	For intLoop = 0 To (objFieldsC.Count - 1)
		rsC.MoveFirst
		if intLoop > 0 Then
			response.write("<tr align='left' bgcolor='#FFE2C6'>")
		end if

		if intLoop = 0 Then
			response.write("<td><strong>RUC</strong></td>")
		end if
		if intLoop = 1 Then
			response.write("<td><strong>Razon Social</strong></td>")
		end if
		if intLoop = 2 Then
			response.write("<td><strong>Grupo Gobierno</strong></td>")
		end if
		if intLoop = 3 Then
			response.write("<td><strong>Actividad Economica</strong></td>")
		end if


	    while not rsC.eof
		    response.write("<td align='left' colspan=2>"&rsC(intLoop)&"</td>")        	
		    rsC.MoveNext
		wend
		response.write("</tr>")
    Next

Set objFieldsR = rsR.Fields
	Set objFieldsD = rsD.Fields

		response.write("<tr bgcolor='#FFE2C6'>")
	For intLoop = 0 To (objFieldsR.Count - 1)
        response.write("<td align='left'><strong>"&objFieldsR.Item(intLoop).Name&"</strong></td>")        	
    Next

	rsC.MoveFirst
    while not rsC.eof
	    For intLoop = 0 To (objFieldsD.Count - 1)
	    	response.write("<td align='left'><strong>"&objFieldsD.Item(intLoop).Name&"</strong></td>")        	
	    Next      	
		rsC.MoveNext
	wend
	
        response.write("</tr>")


	EmpCount=rsC.RecordCount
	ConCount=rsR.RecordCount
    while not rsR.eof

		response.write("<tr >")
		For intLoop = 0 To (objFieldsR.Count - 1)
			if intLoop < 0 then
	        	response.write("<td align='left'>"&rsR(intLoop)&"</td>")        	
	        else
	        	response.write("<td align='left' style='mso-number-format:0.00;'>"&rsR(intLoop)&"</td>")     
	        	'response.write("<td align='right'>"&formatNumber(rsD(intLoop), 2)&"</td>")
	        end if

	    Next


		rsC.MoveFirst
		rsD.Move rsR.AbsolutePosition-1,1
		
		'if rsR.AbsolutePosition>1 then
			'For intLoop = 0 To rsR.AbsolutePosition
				'rsD.MoveNext
		    'Next  	
	    'end if


	    while not rsC.eof
			For intLoop = 0 To (objFieldsD.Count - 1)
	        	response.write("<td align='left' style='vnd.ms-excel.numberformat:#,##0.00;'>"&rsD(intLoop)&"</td>")
	        	'intLoop=EmpCount
		    Next  	
			rsC.MoveNext

			rsD.Move ConCount,0
		wend


	    response.write("</tr>")
		rsR.MoveNext
		'rsD.MoveNext
	wend

	
	rsD.Close
	Set rsD=Nothing
	response.write("</table></td></tr></table>")    

	rsC.Close
	Set rsC=Nothing

	rsR.Close
	Set rsR=Nothing

	Response.ContentType = "application/save" 

%>
