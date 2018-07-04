<!--#include file="Conexion.asp"-->
<html>
<head>
<meta name="description" content="Free Web tutorials" />
<meta name="keywords" content="HTML,CSS,XML,JavaScript" />
<meta name="author" content="Hege Refsnes" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="java_script/inei.css" type="text/css" rel="stylesheet">
<script type="text/javascript" src="java_script/stmenu.js"></script>
<% Response.Expires= 0 
	if Session("tipoAcceso")="" then 
		''c3
		''Response.Redirect("login.html")
	End if
	ruta="imagenes"	
%>
<title>.:INEI-DNCN - Sistema de Consultas de Empresas Públicas</title>
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
    BORDER-RIGHT: #DADBDB 1px solid;
    BORDER-TOP: #DADBDB 1px solid;
    BORDER-LEFT: #DADBDB 1px solid;
    BORDER-BOTTOM: #DADBDB 1px solid;
    BORDER-COLLAPSE: collapse;
    border-spacing: 0;
	font-family:Arial, Geneva, sans-serif;
	font-size:10px;
	width:100%;
	background:#FFFFFF;
	 
	
}
TABLE.tabla1 TD
{
    BORDER-RIGHT: #DADBDB 1px solid;
    BORDER-TOP: #DADBDB 1px solid;
    BORDER-LEFT: #DADBDB 1px solid;
    BORDER-BOTTOM: #DADBDB 1px solid;
	
}
TABLE.tabla1 TH
{
    BORDER-RIGHT: #DADBDB 1px solid;
    PADDING-RIGHT: 5px;
    BORDER-TOP: #DADBDB 1px solid;
    PADDING-LEFT: 5px;
	background:#AD863D;
    PADDING-BOTTOM: 5px;
    BORDER-LEFT: #DADBDB 1px solid;
    PADDING-TOP: 5px;
    BORDER-BOTTOM: #DADBDB 1px solid;
    HEIGHT: 20px;
	color:#000000;
	font-family:Arial, Helvetica, sans-serif;
	font-size:12px;
}

#blocker {
            Z-INDEX: 2000;
            BACKGROUND:#000000; /*: #000; */
            FILTER: alpha(opacity=30);
            LEFT: 0px;
            WIDTH: 100%;
            POSITION: absolute;
            TOP: 0px;
            HEIGHT: 100%;
            opacity: 0.2;
            moz-opacity: 0;
	}

a.a1 {
	font-family: Verdana, Geneva, sans-serif;
	color:#000000;
	font-weight:bold;
	font-size: 8pt; 	
	}
a.a2 {
	font-family: Verdana, Geneva, sans-serif;
	font-size:8pt; 
	font-weight:bold;
	color: #000000;
	}
a.a3 {
	font-family: Verdana, Geneva, sans-serif;
	font-size:8pt; 
	color: #008BC0;
	}
.combo{
	background-color:ffffff;
	border: 1px solid #008BC0;
	cursor: pointer;
	font-size:9pt;
	color: #008BC0;
	}

.combo1 {	background-color:ffffff;
	border: 1px solid #008BC0;
	cursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
.combo2 {	background-color:ffffff;
	border: 1px solid #735F0D;
	cursor: pointer;
	font-size:9pt;
	color: #735F0D;
}

.combo21 {background-color:ffffff;
	border: 1px solid #735F0D;
	cursor: pointer;
	font-size:9pt;
	color: #735F0D;
}
</style>
<body  onLoad="javascript:CargaSInst();">
<div id="blocker" name="blocker" style="display:none;" >
<table width="100%" height="100%" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><center><img src="Imagenes/progressbar.gif"></img></center></td>
  </tr>
</table>
</div>
<table width="100%" height="72" border="0" cellpadding="0" cellspacing="0" background="<%=ruta%>/sunat_fondo.jpg">
  <tr>
    <td>
  	  <table width="359" height="72" background="<%=ruta%>/sunat_izq.jpg" bgcolor="557163">
    	<tr>
		  <td></td>
		</tr>
      </table>
    </td>
    <td align="right">
  	  <table width="333" height="72" background="<%=ruta%>/sunat_der.jpg">
        <tr>
		  <td><!--Usuario: <%Response.write Session("id_usuario") %>--></td>
		</tr>
      </table>
    </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="10">
	  <script type="text/javascript" language="JavaScript1.2" src="java_script/menu.js"></script>
	  <a href="logoffce.asp" style="font-family:verdana; font-size:10px; color:#000000"></a>
	</td>
  </tr>          
</table>

<%
	anexo=Request.QueryString("anexo")
%>

<STRONG>
<SCRIPT language=JavaScript1.2 type=text/javascript src="java_script/funciones.js"></SCRIPT>
<BR><FONT color=#000000 size=3 face=Arial>Consistencias Básicas </font></strong>-Anexo <%=anexo%>- En Nuevos Soles</FONT></STRONG>

<BR><BR>
<form action="FrmConsistenciasBasicas.asp" name="fmrEF" method="post" target=_self>

<a class=a2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Año</a>
<select name="cboAnio" id="cboAnio" onchange="CargaSInst(this.value,1)">

<% 
    Set rs = Server.CreateObject("ADODB.Recordset")
	rs.CursorLocation=3
	rs.Open "exec sp_EstadosFinancieros_ListarAnio" , con

	chck="selected='selected'"

	while not rs.eof	
%>
	<option value='<%=Trim(rs("ano_eje"))%>' <%=chck%> ><%=Trim(rs("ano_eje"))%></option>
<%
	chck=""
    rs.MoveNext
	wend
	rs.Close
    Set rs = Nothing
%>
</select>

<A class=a2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Sector</A>

	<select name="cboSector" id="cboSector">
	 <option value="S111">S111</option>
    </select>

<A class=a2>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</A>
<%
		'strOnchange2="cargaVariable2(1,'IFECON');"
		'anexo=Request.QueryString("anexo")
		strOnchange2="cargaAnexo("&anexo&"); return false;"
		strOnchange3="cargaExcelAnexo("&anexo&"); return false;"
	%>
	<button onClick="javascript:<%=strOnchange2%>" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
		<img  src="Imagenes/search.png" width="20" height="20" alt="Buscar Consulta" >
	</button>&nbsp;&nbsp;
	<button onClick="javascript:<%=strOnchange3%>" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
		<img  src="imagenes/excel.png" width="20" height="20" alt="Exportar a Excel" >
	</button>&nbsp;&nbsp;
</td>
<BR><BR>
<div id="DivVariables" style="overflow:scroll;height:420px; width:100%"></div>

</FORM>
</body>
</html> 