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
	end if
	
	ruta="imagenes"
	titulo="ANEXO 6 - DEPRECIACIÓN ACUMULADA DE INMUEBLES, MAQUINARIA Y EQUIPO"
	pagina="Anexo6"
	tabla="anexo6_res"
%>

<script type="text/javascript">

	$(document).ready(function() {
	  $.ajaxSetup({ cache: false });
	});

	function InicializaFiltros()
	{	CargarAnio();	}

	function CargarAnio()
	{
		CargaFiltro("cboAnio",'Filtros.asp?rep=anioAnexo&data="<%=tabla%>"','cargaGContable');
	}

	function cargaSectorSicon(){
		var Anio = document.getElementById("cboAnio").value;
		CargaFiltro("cboSecSicon",'Filtros.asp?rep=SecSicon_RepAnio&anio='+Anio+'&eeff="01"','cargaCodigos');
	}

	function cargaGContable(){
		var Anio = document.getElementById("cboAnio").value;
		CargaFiltro("cboGruCon",'Filtros.asp?rep=gcont_Anexo&data="<%=tabla%>"&anio='+Anio+'','cargaCodigos');
	}

	function cargaCodigos(){
		var nivel = document.getElementById("cboNivel").value;
		if(nivel==0){
			var c = document.getElementById("cboCodigo");
			c.innerHTML = "";
			var option = document.createElement("option");
			option.text = "--";			
			option.value = "0";
			c.add(option);
		}
		else{
			var Anio = document.getElementById("cboAnio").value;
			var GruCont = document.getElementById("cboGruCon").value;
			
			CargaFiltro("cboCodigo",'Filtros.asp?rep=codigo_Anexo&data="<%=tabla%>"&anio='+Anio+'&gcon='+GruCont+'&niv='+nivel,'');
		}
	}

	function switchCodigo(){
		var x = document.getElementById("cboCodigo");
		var c = document.getElementById("cboNivel");
		var det = document.getElementById("cboDetalle").value;

		if(det==1){//Agrupado
			x.multiple=true;
			x.size=5;
			x.width=30;
			$("#cboNivel option[value='0']").remove();
		}
		else//Por Empresa
		{
			x.multiple=false;
			x.size=0;
			x.width=60;
			var option = document.createElement("option");
			option.text = "TODOS";
			option.value =  "0";
			c.add(option);
		}
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
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}

.combo3 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
.combo4 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
.combo5 {	background-color:ffffff;
	border: 1px solid #008BC0;
	ursor: pointer;
	font-size:9pt;
	color: #008BC0;
}
</style>
<body onLoad="InicializaFiltros();">
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

<div align="center"><strong>
  <br>
  <font face="Arial" size='3pt' color='#00 0000'><%=titulo%></font></strong> <br />
</div>
<form action="<%=pagina%>.asp" name="fmrEF" method="post" target=_self>

 	<a class="a2">&nbsp;&nbsp;Periodo</a>
 	<select name="cboAnio" id="cboAnio" class="combo2" onChange="cargaGContable();"></select>

	<A id="detalle" class=a2>&nbsp;&nbsp;Detalle</A>
	<select name="cboDetalle" id="cboDetalle" class="combo2" onChange="switchCodigo();">
		<option value='0'>Por empresa</option>
		<option value='1'>Agrupado</option>
	</select>
	
	<A class=a2>&nbsp;&nbsp;Tipo</A>
	<select name="cboGruCon" id="cboGruCon" style="width:200px" class="combo2" onChange="cargaCodigos();">
	</select>

	<A class=a2>&nbsp;&nbsp;Nivel</A>
	<select name="cboNivel" id="cboNivel" class="combo2" onChange="cargaCodigos();">
        <option value="5" selected="selected">Nv AE 101</option>
        <option value="10" >Nv AE 54</option>
        <option value="11" >Nv AE 14</option>
        <option value="12" >SECTOR SICON</option>
        <option value="0" >TODOS</option>
    </select>

	<select name="cboCodigo" id="cboCodigo" class="combo2" style="width:320px">
    </select>
	
	<select name="cboMoneda" id="cboMoneda" class="combo2">
        <option value="0" selected="selected">Soles</option>
        <option value="1" >Miles de Soles</option>
        <option value="2" >Millones de Soles</option>
    </select>

    <a class="a2">&nbsp;&nbsp;</a>
	<button onClick="cargaVariableEF('<%=pagina%>'); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/search.png) no-repeat;" alt="Buscar Consulta"></button>
	<button onClick="ExcelEF('<%=pagina%>'); return false;" style="border:none;height:21px; width:21px;background: url(imagenes/excel.png) no-repeat;" alt="Exportar a Excel"></button>
<br>

<!--<div id="DivVariables" style="overflow:auto;height='400 px'; width='100%'"></div>-->
<div id="DivVariables" style="overflow:scroll;height:420px; width:100%"></div>
  
</FORM>
</body>
</html>

