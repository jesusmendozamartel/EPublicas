<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<!--#include file="Conexion.asp"-->
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="description" content="Free Web tutorials" />
<meta name="keywords" content="HTML,CSS,XML,JavaScript" />
<meta name="author" content="Hege Refsnes" />
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />

<link href="css/pivot.css" type="text/css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="css/bootstrap/css/bootstrap.min.css">
<link href="css/inei.css" type="text/css" rel="stylesheet">

<script type="text/javascript" src="java_script/stmenu.js"></script>
<script type="text/javascript" src="java_script/linq.js"></script>
<script type="text/javascript" src="java_script/jquery.min.js"></script>
<script type="text/javascript" src="java_script/jquery-ui.min.js"></script>
<script type="text/javascript" src="java_script/jquery.ui.touch-punch.min.js"></script>
<script type="text/javascript" src="css/bootstrap/js/bootstrap.js"></script>

<!-- BEGIN PIVOT --->
<script type="text/javascript" src="java_script/pivot/export.js"></script>
<script type="text/javascript" src="java_script/pivot/pivot.js"></script>
<script type="text/javascript" src="java_script/pivot/pivot.es.js"></script>
<script type="text/javascript" src="java_script/pivot/plotly-basic-latest.min.js"></script>
<script type="text/javascript" src="java_script/pivot/plotly_renderers.js"></script>
<!-- END PIVOT --->

<% Response.expires = 0 
	if Session("tipoAcceso")="" then 
		'Response.redirect "login.html"
	end if
	ruta="imagenes"
	''Response.Write Session("id_usuario")
%>

<title>.:INEI-DNCN - Sistema de Consultas de Empresas Públicas </title>
<body>
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

<div class="panel panel-primary">
  <div class="panel-heading">
    <h2 align="center" class="panel-title" style="padding:10px;">Tabla Dinámica: Dividendos declarados</h2>
  </div>
  <div class="panel-body">
    <div style="padding-left:30px">

	<div class="btn-group btn-group-sm">
	<button type="button" class="btn btn-success" onClick="Exporter.export(pvtTable, 'tabla_pivot.xls', 'exportado');return false;">
      <span class="glyphicon glyphicon-export"></span> Exportar
    </button>
	</div>
	
 </div>
	<div id="output" style="margin: 30px;" align="center" ></div>
  </div>
  <div class="panel-footer">DNCN - Dirección Nacional de Cuentas Nacionales </div>
</div>
<script type="text/javascript">
	$(function(){
	var tpl = $.pivotUtilities.aggregatorTemplates;    
	var derivers = $.pivotUtilities.derivers;
	var renderers = $.extend($.pivotUtilities.renderers,
	$.pivotUtilities.plotly_renderers);

	$.getJSON("data/dividendos.json", function(mps) {
		$("#output").pivotUI(mps, {
            renderers: renderers,
            hiddenAttributes: ["TOTAL"],  
			aggregators: {
				"Total":  
					function() {return tpl.sum()(["TOTAL"])}
			},
			rows: ["CODIGO SICON","ENTIDAD"], cols: ["AÑO"],
		},true,"es");
	});
    });

    $(document).ready(function() {
        $(".pvtTotalLabel.pvtTotalColSortable").addClass('oculto');
        $(".pvtTotal.colTotal").addClass('oculto');
    });

</script>
</body>
</html>

