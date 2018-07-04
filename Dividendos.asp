<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Conexion.asp"-->
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="description" content="Free Web tutorials" />
<meta name="keywords" content="HTML,CSS,XML,JavaScript" />
<meta name="author" content="Hege Refsnes" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<link href="css/inei.css" type="text/css" rel="stylesheet">

<script type="text/javascript" src="java_script/stmenu.js"></script>
<script type="text/javascript" src="java_script/linq.js"></script>
<script type="text/javascript" src="java_script/jquery-3.1.1.min.js"></script>
<script type="text/javascript" language="JavaScript1.2" src="java_script/funciones.js"></script>

<% Response.expires = 0 
	if Session("tipoAcceso")="" then 
		'Response.redirect "login.html"
	end if
	ruta="imagenes"
	''Response.Write Session("id_usuario")
%>

<script type="text/javascript">

	$(document).ready(function() {
	  $.ajaxSetup({ cache: false });
	});

	function InicializaFiltros()
	{	CargarAnio();	}

	function CargarAnio()
	{
		CargaFiltro("cboAnio",'Filtros.asp?rep=anio','');
	}

</script>

<title>.:INEI-DNCN - Sistema de Consultas de Empresas Públicas </title>
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
	color: #ffffff;
	}
a.a3 {
	font-family: Verdana, Geneva, sans-serif;
	font-size:8pt; 
	color: #008BC0;
	}
.combo{
	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
	}

.combo1 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
.combo2 {	background-color:ffffff;
	border: 1px solid #735F0D;
	ursor: pointer;
	font-size:9pt;
	color: #735F0D;
}

.combo21 {background-color:ffffff;
	border: 1px solid #735F0D;
	ursor: pointer;
	font-size:9pt;
	color: #735F0D;
}
</style>
<body onload="InicializaFiltros();">
<div id="blocker" name="blocker" style="display:none;" ><table width="100%" height="100%" border="1" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><center><img src="Imagenes/progressbar.gif"></img></center></td>
  </tr>
</table>
</div>
<table width="100%" height="72" border="0" cellpadding="0" cellspacing="0" background="<%=ruta%>/sunat_fondo.jpg">
  <tr>
  <td>
  	<table width="359" height="72" background="<%=ruta%>/sunat_izq.jpg" bgcolor="557163">
    	<tr><td></td>
    	</tr>
    </table>  </td>
  <td align="right">
  	  <table width="333" height="72" background="<%=ruta%>/sunat_der.jpg">
    	<tr><td></td></tr>
    </table>  </td>
  </tr>
</table>

<table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td height="10"><strong>
      <script type="text/javascript" language="JavaScript1.2" src="java_script/menu.js"></script>
    </strong></td>
  </tr>          
</table>

<strong>
<br>
<font face="Arial" size='3pt' color='#000000'>Dividendos Declarados</font></strong>
<form action="Dividendos.asp" name="fmrEF" method="post" target=_self>

 	<a class="a1">&nbsp;&nbsp;Año</a>
 	<select name="cboAnio" id="cboAnio" class="combo2"></select>

    <a class="a2">&nbsp;&nbsp;</a>
	<button onClick="cargaDividendos(); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/search.png) no-repeat;" alt="Buscar Consulta"></button>
	<button onClick="cargaDividendosExcel(); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/excel.png) no-repeat;" alt="Exportar a Excel"></button>
<br>

<!--<div id="DivVariables" style="overflow:auto;height='400 px'; width='100%'"></div>-->
<div id="DivVariables" style="overflow:scroll;height:420px; width:100%"></div>
  
</FORM>
</body>
</html>

