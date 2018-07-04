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
<title>.:INEI-DNCN - Sistema de Consultas de la SUNAT</title>
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
<body >
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

<div align="center">
  <script type="text/javascript" language="JavaScript1.2" src="java_script/funciones.js"></script>
  <br />
  <strong><font face="Arial" size='3pt' color='#000000'>Información Financiera a nivel de Actividad Económica </font></strong>- En Nuevos Soles</div>
<form action="InfFinancieraactecon.asp" name="fmrEF" method="post" target=_self>
<%		strUbicaGrupo=Request("cboGrupo")
		strUbicaAnio=Request("cboAnio")
		strOnchange="ActualizarEF();"		
%>

  <div name="formulario" id="formulario1">
    <a class="a1">Seleccione:</a> 
	<a class="a1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Año:</a>
	<select name="cboAnio" id="cboAnio" class="combo2" >
      <option value="" selected>
        <%Response.Write "[Seleccione]"%>
      </option>
      <%  SQL = "select des_annio from Sunatma_periodo order by 1"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open SQL , con
			If Not rs.BOF Then rs.MoveFirst
			Do While Not rs.EOF			
		%>
      <option value="<%=Trim(rs("des_annio"))%>"<%If Trim(CStr(strUbicaAnio)) = Trim(CStr(rs("des_annio"))) Then%>selected<% End If%>><%=Trim(rs("des_annio"))%> </option>
      <%  rs.MoveNext
			Loop
			rs.Close
	       	Set rs = Nothing
		   	SQL=""
		%>
    </select>

	<a class="a1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Grupo de Información:</a>
	<select name="cboGrupo" id="cboGrupo" class="combo2" >
	  <option value="" selected><%Response.Write "[Seleccione]"%></option>
		<%  
		SQL = "SELECT DISTINCT GRUPO FROM Sunatpro_InfFinancieraACTECON WHERE GRUPO <> 'G5' ORDER BY 1"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open SQL , con
			If Not rs.BOF Then rs.MoveFirst
			Do While Not rs.EOF			
		%>
	  <option value="<%=Trim(rs("GRUPO"))%>"<%If Trim(CStr(strUbicaGrupo)) = Trim(CStr(rs("GRUPO"))) Then%>selected<% End If%>><%=Trim(rs("GRUPO"))%> </option>
		<%  rs.MoveNext
			Loop
			rs.Close
	       	Set rs = Nothing
		   	SQL=""
		%>
    </select>
	<a class="a1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Nivel:</a>
	<select name="cboNivel" id="cboNivel" class="combo2" >
	  <option value="" selected><%Response.Write "[Seleccione]"%></option>
		<%  
		SQL = "SELECT * FROM Sunatma_nivel ORDER BY 1"
			Set rs = Server.CreateObject("ADODB.Recordset")
			rs.Open SQL , con
			If Not rs.BOF Then rs.MoveFirst
			Do While Not rs.EOF			
		%>
	  <option value="<%=Trim(rs("cod_nivel"))%>"<%If Trim(CStr(strUbicaGrupo)) = Trim(CStr(rs("cod_nivel"))) Then%>selected<% End If%>><%=Trim(rs("desc_nivel"))%> </option>
		<%  rs.MoveNext
			Loop
			rs.Close
	       	Set rs = Nothing
		   	SQL=""
		%>
    </select>

		&nbsp;&nbsp;
		<%
			strOnchange2="cargaVariable2(1,'IFECON');"
			strOnchange3="Excel(1,'IFECON');"
		%>
		<button onClick="javascript:<%=strOnchange2%>" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
			<img  src="Imagenes/search.png" width="20" height="20" alt="Buscar Consulta" >
		</button>&nbsp;&nbsp;
		<button onClick="javascript:<%=strOnchange3%>" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
			<img  src="imagenes/excel.png" width="20" height="20" alt="Exportar a Excel" >
		</button>&nbsp;&nbsp;
		<button onClick="Refresh(4)" style=" border:none; height:21px; width:21Px;font-weight:bold;font-size:8pt;background-color:#ffffff;color:#123456;">
			<img  src="imagenes/refresh.png" width="20" height="20" alt="Refrescar" >
		</button>
		
  </div>

<div id="DivVariables" style="overflow:auto;height='420 px'; width='100%'"></div>

  <input type="hidden" name="hidGrupo" id="hidGrupo" value="<%=strUbicaGrupo%>">
  <input type="hidden" name="hidAnio"  id="hidAnio" value="<%=strUbicaAnio%>">
</FORM>
</body>
</html> 