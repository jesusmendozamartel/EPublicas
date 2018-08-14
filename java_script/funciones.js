function Ajax() 
{
  var xmlHttp=null;
  if (window.ActiveXObject) 
    xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
  else 
    if (window.XMLHttpRequest) 
      xmlHttp = new XMLHttpRequest();
  return xmlHttp;
}

var conexion2;	
//
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ActualizarEF(){
  	if (document.fmrEF.cboReporte.value!="")		
		{reporte = document.fmrEF.cboReporte.value;}
	if(document.fmrEF.cboReporte.value=="")
		{reporte="";}
	document.fmrEF.hidReporte.value=reporte;
	document.fmrEF.submit();
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ActualizarAnio(){
  	if (document.fmrEF.hidReporte.value!="")		
		{reporte = document.fmrEF.hidReporte.value;}
	if(document.fmrEF.hidReporte.value=="")
		{reporte="";}
	document.fmrEF.hidReporte.value=reporte;
	document.fmrEF.submit();
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaVariableEF(estadofinanciero)
{
	//alert(estadofinanciero);return false;
	
	var xannio=document.getElementById('cboAnio').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xGrupoCont=document.getElementById('cboGruCon').value;
	var xdet=document.getElementById('cboDetalle').value;
	var moneda=document.getElementById('cboMoneda').value;

	if (xannio==''){
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	if (xcod==''){
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}
	
	conexion2=Ajax();
	muestra();
	var url= ''

	var url= estadofinanciero+'_Reporte.asp';	
	url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&gcont='+xGrupoCont+'&detalle='+xdet+'&moneda='+moneda;
	
	conexion2.open('POST',url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaDetCamPat(CodiEnt)
{
	var xannio=document.getElementById('cboAnio').value;
	
	conexion2=Ajax();
	muestra();
	
	var url= 'CambioPatrimonio_Detalle.asp';
	url=url+'?annio='+xannio+'&CodiEnt='+CodiEnt

	//alert(url);
	conexion2.open('POST',url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaDetCamPatExcel(CodiEnt)
{
	var xannio=document.getElementById('cboAnio').value;
	
	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	var url= 'CambioPatrimonio_DetalleExcel.asp';
	url=url+'?annio='+xannio+'&CodiEnt='+CodiEnt

	document.fmrEF.action=url;
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaDirectorio()
{
	var xannio=document.getElementById('cboAnio').value;
	
	if (xannio==''){
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	var url= 'Directorio_Reporte.asp';

	url=url+'?annio='+xannio;

	conexion2.open('POST',url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaDirectorioExcel()
{
	var xannio=document.getElementById('cboAnio').value;
	
	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	var url= 'Directorio_ReporteExcel.asp';
	url=url+'?annio='+xannio;

	document.fmrEF.action=url;
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaBC()
{
	var xannio=document.getElementById('cboAnio').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xGrupoCont=document.getElementById('cboGruCon').value;
	//var xSecSic=document.getElementById('cboSecSicon').value;
	var xdet=document.getElementById('cboDetalle').value;

	if (xannio==''){
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	if (xcod==''){
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= 'BalanceComprobacion_Reporte.asp';

	url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&gcont='+xGrupoCont+'&detalle='+xdet;
	//url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&secsic='+xSecSic+'&detalle='+xdet;
	
	//alert(url);
	conexion2.open('POST',url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaCSI()
{
	var xannio=document.getElementById('cboAnio').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xGrupoCont=document.getElementById('cboGruCon').value;
	//var xSecSic=document.getElementById('cboSecSicon').value;
	var xdet=document.getElementById('cboDetalle').value;

	if (xannio==''){
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	if (xcod==''){
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= 'BalanceCSI_Reporte.asp';

	url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&gcont='+xGrupoCont+'&detalle='+xdet;
	//url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&secsic='+xSecSic+'&detalle='+xdet;
	
	//alert(url);
	conexion2.open('POST',url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);		
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaVariable2(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
	var xgrupo=document.getElementById('cboGrupo').value;
	var xnivel=document.getElementById('cboNivel').value;
	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xgrupo=='')  {
			alert ('Seleccione el Grupo');
			document.getElementById('cboGrupo').focus();
			return false;
		}	
		if (xnivel=='')  {
			alert ('Seleccione el Nivel');
			document.getElementById('cboNivel').focus();
			return false;
		}	

		conexion2=Ajax();
		muestra();
		if(estadofinanciero=='IFECON')
		{var url= 'InfFinancieraactecon_Reporte.asp';}
		if(estadofinanciero=='PO')
		{var url= 'PersonalOcupado_Reporte.asp';}

		url=url+'?annio='+xannio+'&grupo='+xgrupo+'&nivel='+xnivel;
		//alert(url);
		conexion2.open('GET', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaSistemaInter(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
	
		conexion2=Ajax();
		muestra();
		if(estadofinanciero=='SI')
		{var url= 'SistemaIntermedio_Reporte.asp';}
		if(estadofinanciero=='IN')
		{var url= 'Inversion_Reporte.asp';}
		if(estadofinanciero=='PNM')
		{var url= 'ProduccionNMetalica_Reporte.asp';}
		if(estadofinanciero=='PO')
		{var url= 'PersonalOcupadoCon_Reporte.asp';}
		
		url=url+'?annio='+xannio;
		
		conexion2.open('GET', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaSistemaInter2(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
	var xup=document.getElementById('cboUP').value;
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
	
		conexion2=Ajax();
		muestra();
		if(estadofinanciero=='IN')
		{var url= 'Inversion_Reporte.asp';}
		
		url=url+'?annio='+xannio+'&up='+xup;
		
		conexion2.open('GET', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaSistemaInter3(valor,estadofinanciero){
	if (valor=='1') {
		var xannio=document.getElementById('cboAnio').value;
		var xup=document.getElementById('cboUP').value;
		var xtipo=document.getElementById('cboTipo').value;
		
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
	
		conexion2=Ajax();
		muestra();
		if(estadofinanciero=='PMM')
		{var url= 'ProduccionMMetalica_Reporte.asp';}
		
		url=url+'?annio='+xannio+'&up='+xup+'&tipo='+xtipo;
		conexion2.open('GET', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
	}
	if (valor=='2') {
		var xannio=document.getElementById('cboAnio').value;
		var xtipo=document.getElementById('cboTipo').value;
		
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
	
		conexion2=Ajax();
		muestra();
		if(estadofinanciero=='PMC')
		{var url= 'ProduccionMCarbonifera_Reporte.asp';}
		
		url=url+'?annio='+xannio+'&tipo='+xtipo;
		conexion2.open('GET', url, true);
		conexion2.setRequestHeader('Content-Type', 'text/html');
		conexion2.setRequestHeader('encoding', 'iso-8859-1');
		conexion2.onreadystatechange = procesaVariables;
		conexion2.send(null);
	}
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function procesaVariables(){

   if (conexion2.readyState == 4){
	   if(conexion2.responseText == 1 || conexion2.responseText == 0 ){
		  	alert("No se encontraron registros!");
		}else{
   			//alert(conexion2.responseText);
			document.getElementById('DivVariables').innerHTML = conexion2.responseText;
		}
	oculta();
  } 

}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function muestra(){
	//alert("muestra");
	document.getElementById('blocker').style.display="block";
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function oculta(){
	//	alert("oculta");
	document.getElementById('blocker').style.display="none";
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function Activa(valor){
	if (valor=="M")
	{
	 document.getElementById("EmEs2").disabled= true;
	 document.getElementById("Aju2").disabled= true;
	}
	else
	{
	 document.getElementById("EmEs2").disabled= false;
	 document.getElementById("Aju2").disabled= false;
	}
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function Procesar()
{
	var xannio=documeCargarAniont.getElementById("cboAnio").value;
	var xtipo=document.getElementById("hidTipo").value;
	var url="Procesar_Reporte.asp"+"?annio="+xannio+"&tipo="+xtipo;

	conexion2=Ajax();
	muestra();
	conexion2.open("GET", url, true);
	conexion2.setRequestHeader("Content-Type", "text/html");
	conexion2.setRequestHeader("encoding", "iso-8859-1");
	conexion2.onreadystatechange = procesaVariables;

	conexion2.send(null);
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ExcelEF(estadofinanciero)
{
	var xannio=document.getElementById('cboAnio').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xGrupoCont=document.getElementById('cboGruCon').value;
	var xdet=document.getElementById('cboDetalle').value;
	var cboCod=document.getElementById('cboCodigo');
	var cboDet=document.getElementById('cboDetalle');
	var moneda=document.getElementById('cboMoneda').value;
	
	var xcodText= cboCod.options[cboCod.selectedIndex].text;
	var xdetText= cboDet.options[cboDet.selectedIndex].text;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}

	var url= estadofinanciero+'_ReporteExcel.asp';

	url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&gcont='+xGrupoCont+'&detalle='+xdet+'&codText='+xcodText+'&detText='+xdetText+'&eeff='+estadofinanciero+'&moneda='+moneda;
	//alert(url);
	document.fmrEF.action=url;
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ExcelBC()
{
	var xannio=document.getElementById('cboAnio').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xGrupoCont=document.getElementById('cboGruCon').value;
	//var xSecSic=document.getElementById('cboSecSicon').value;
	var xdet=document.getElementById('cboDetalle').value;
	var cboCod=document.getElementById('cboCodigo');
	var cboDet=document.getElementById('cboDetalle');

	var xcodText= cboCod.options[cboCod.selectedIndex].text;
	var xdetText= cboDet.options[cboDet.selectedIndex].text;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}

	var url= 'BalanceComprobacion_ReporteExcel.asp';

	url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&gcont='+xGrupoCont+'&detalle='+xdet+'&codText='+xcodText+'&detText='+xdetText;
	//alert(url);
	document.fmrEF.action=url;
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ExcelCSI()
{
	var xannio=document.getElementById('cboAnio').value;
	var xNiv=document.getElementById('cboNivel').value;
	var xcod=Enumerable.From($('#cboCodigo option:selected')).Select(function (x) { return x.value }).ToArray().join();
	var xGrupoCont=document.getElementById('cboGruCon').value;
	//var xSecSic=document.getElementById('cboSecSicon').value;
	var xdet=document.getElementById('cboDetalle').value;
	var cboCod=document.getElementById('cboCodigo');
	var cboDet=document.getElementById('cboDetalle');

	var xcodText= cboCod.options[cboCod.selectedIndex].text;
	var xdetText= cboDet.options[cboDet.selectedIndex].text;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	if (xcod=='')  {
		alert ('Seleccione Sector');			
		document.getElementById('cboCodigo').focus();
		return false;
	}

	var url= 'BalanceCSI_ReporteExcel.asp';

	url=url+'?annio='+xannio+'&nivel='+xNiv+'&codigo='+xcod+'&gcont='+xGrupoCont+'&detalle='+xdet+'&codText='+xcodText+'&detText='+xdetText;
	//alert(url);
	document.fmrEF.action=url;
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ExcelSistemaInter(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if(estadofinanciero=='SI'){
		document.fmrEF.action='SistemaIntermedio_ReporteExcel.asp'+"?annio="+xannio;
		document.fmrEF.submit();
		document.fmrEF.action="SistemaIntermedio.asp";
		document.fmrEF.target='_self';
		}
		if(estadofinanciero=='IN'){
		document.fmrEF.action='Inversion_ReporteExcel.asp'+"?annio="+xannio;
		document.fmrEF.submit();
		document.fmrEF.action="Inversion.asp";
		document.fmrEF.target='_self';
		}
		if(estadofinanciero=='PNM'){
		document.fmrEF.action='ProduccionNMetalica_ReporteExcel.asp'+"?annio="+xannio;
		document.fmrEF.submit();
		document.fmrEF.action="ProduccionNMetalica.asp";
		document.fmrEF.target='_self';
		}
		if(estadofinanciero=='PO'){
		document.fmrEF.action='PersonalOcupadoCon_ReporteExcel.asp'+"?annio="+xannio;
		document.fmrEF.submit();
		document.fmrEF.action="PersonalOcupadoCon_Reporte.asp";
		document.fmrEF.target='_self';
		}

}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ExcelSistemaInter2(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
	var xup=document.getElementById('cboUP').value;

	
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}

				

		if(estadofinanciero=='IN'){
		document.fmrEF.action='Inversion_ReporteExcel.asp'+"?annio="+xannio+'&up='+xup;
		document.fmrEF.submit();
		document.fmrEF.action="Inversion_Reporte.asp";
		document.fmrEF.target='_self';
		}


}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ExcelSistemaInter3(valor,estadofinanciero){
if (valor=='1'){
	var xannio=document.getElementById('cboAnio').value;
	var xup=document.getElementById('cboUP').value;
	var xtipo=document.getElementById('cboTipo').value;
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}

		if(estadofinanciero=='PMM'){
		document.fmrEF.action='ProduccionMMetalica_ReporteExcel.asp'+"?annio="+xannio+'&up='+xup+'&tipo='+xtipo;
		document.fmrEF.submit();
		document.fmrEF.action="ProduccionMMetalica_Reporte.asp";
		document.fmrEF.target='_self';
		}
}
if (valor=='2'){
	var xannio=document.getElementById('cboAnio').value;
	var xtipo=document.getElementById('cboTipo').value;
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}

		if(estadofinanciero=='PMC'){
		document.fmrEF.action='ProduccionMCarbonifera_ReporteExcel.asp'+"?annio="+xannio+'&tipo='+xtipo;
		document.fmrEF.submit();
		document.fmrEF.action="ProduccionMCarbonifera_Reporte.asp";
		document.fmrEF.target='_self';
		}
}

}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function Excel(valor,estadofinanciero){
	var xannio=document.getElementById('cboAnio').value;
	var xgrupo=document.getElementById('cboGrupo').value;
		if (xannio=='')  {
			alert ('Seleccione Año');
			document.getElementById('cboAnio').focus();
			return false;
		}
		if (xgrupo=='')  {
			alert ('Seleccione el Grupo');
			document.getElementById('cboGrupo').focus();
			return false;
		}	

		if(estadofinanciero=='DI'){
		document.fmrEF.action='Directorio_ReporteExcel.asp'+"?annio="+xannio+"&reporte="+xreporte;
		document.fmrEF.submit();
		document.fmrEF.action="Directorio.asp";
		document.fmrEF.target='_self';
		}
		if (estadofinanciero=='IF') { 
		document.fmrEF.action='InfFinancieraciiu_ReporteExcel.asp'+"?annio="+xannio+"&grupo="+xgrupo;
		document.fmrEF.submit();
		document.fmrEF.action="InfFinancieraciiu.asp";
		document.fmrEF.target='_self';
		}
		if (estadofinanciero=='IFECON') { 
		document.fmrEF.action='InfFinancieraactecon_ReporteExcel.asp'+"?annio="+xannio+"&grupo="+xgrupo;
		document.fmrEF.submit();
		document.fmrEF.action="InfFinancieraactecon.asp";
		document.fmrEF.target='_self';
		}
		if (estadofinanciero=='IFSI') { 
		document.fmrEF.action='InfFinancierasi_ReporteExcel.asp'+"?annio="+xannio+"&grupo="+xgrupo;
		document.fmrEF.submit();
		document.fmrEF.action="InfFinancierasi.asp";
		document.fmrEF.target='_self';
		}

}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function alerta()
{
	alert ("No hay ningun registro");
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function Refresh(valor)
{
	
	if (valor==1) {document.URL = "Directorio.asp";}
	if (valor==2) {document.URL = "InfFinancieraciiu.asp";}
	if (valor==3) {document.URL = "InfFinancierasi.asp";}
	if (valor==4) {document.URL = "InfFinancieraactecon.asp";}
	if (valor==5) {document.URL = "Inversion.asp";}
	if (valor==6) {document.URL = "ProduccionMMetalica.asp";}
	if (valor==7) {document.URL = "ProduccionNMetalica.asp";}
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
//Funciones C3
function cargaEstFin(){
	
	var strAnio="";
	var strFor="";
	var strGru="";
	var strSec="";
	var strForText="";

	//Check Multple - C3
	/*var x = document.getElementsByName("anios");
	var i;
	for (i = x.length-1; i >=0 ; i--) {
        if (x[i].checked ==true) {
        	strAnio+=","+x[i].value;
    	}
	}	

	var x = document.getElementsByName("formato");
	var i;
	for (i = x.length-1; i >=0 ; i--) {
        if (x[i].checked ==true) {
        	strFor+=","+x[i].value;
    	}
	}

	var x = document.getElementsByName("sect");
	var i;
	for (i = 0; i <x.length;  i++) {
        if (x[i].checked ==true) {
        	strSec+=","+x[i].value;
    	}
	}*/

	var elem = document.getElementById("cboGruCont");
	if (elem === null) alert('Seleccione un Sector y a continuación un grupo.');
	else
	{

		var strAnio = document.getElementById("cboAnio").value;
		var strFor = document.getElementById("cboFormato").value;
		var strSec = document.getElementById("cboSecInt").value;
		var strGru = document.getElementById("cboGruCont").value;

	    var elt = document.getElementById("cboFormato");

	    var strForText= elt.options[elt.selectedIndex].text;

		//strAnio=strAnio.substring(1, strAnio.length);
		//strFor=strFor.substring(1, strFor.length);
		//strSec=strSec.substring(1, strSec.length);

		if (strGru==""){
			alert("Seleccione al menos un grupo");
		}
		else
		{
			conexion2=Ajax();
			muestra();
			var url= 'InfFinancieraactecon_Reporte.asp'+"?anio="+strAnio+"&for="+strFor+"&sec="+strSec+"&ForText="+strForText+"&strGru="+strGru;
			
			conexion2.open('GET', url, true);
			conexion2.setRequestHeader('Content-Type', 'text/html');
			conexion2.setRequestHeader('encoding', 'iso-8859-1');
			conexion2.onreadystatechange = procesaVariables;

			conexion2.send(null);
		}	
	}
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaAnexo(anexo){
	
	var strAnio="";
	var strSec="";

	var strAnio = document.getElementById("cboAnio").value;
	var strSec = document.getElementById("cboSector").value;
	conexion2=Ajax();
	muestra();
	var url= 'FrmConsistenciasBasicas_Reporte.asp'+"?anio="+strAnio+"&ane="+anexo+"&sec="+strSec;
	
	conexion2.open('GET', url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;

	conexion2.send(null);
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaExcelAnexo(anexo){
	
	var strAnio="";
	var strSec="";

	var strAnio = document.getElementById("cboAnio").value;
	var strSec = document.getElementById("cboSector").value;
	conexion2=Ajax();

	document.fmrEF.action='FrmConsistenciasBasicas_ReporteExcel.asp'+"?anio="+strAnio+"&ane="+anexo+"&sec="+strSec;
	//document.fmrEF.action='MultipleExcelSheet.asp';
	
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function CargaGruCont(){
		
	var anio =document.getElementById("cboAnio").value;
	var sinst =document.getElementById("cboSecInt").value;
	
	conexion2=Ajax();
	muestra();

	var url= 'Filtros.asp'+"?anio="+anio+"&SecInst="+sinst+"&filtro="+2;

	conexion2.open('GET', url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables2;

	conexion2.send(null);
	document.getElementById('gcontableT').style.color ="#000000";
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function CargaSInst(){
	var anio =document.getElementById("cboAnio").value;
	var tipsoc =document.getElementById("cboSociedad").value;
	
	conexion2=Ajax();
	muestra();

	var url= 'Filtros.asp'+"?anio="+anio+"&filtro="+1+"&TipSoc="+tipsoc;

	conexion2.open('GET', url, true);
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables3;

	conexion2.send(null);	
	
	document.getElementById('gcontable').innerHTML = "";
	document.getElementById('gcontableT').style.color ="#ffffff";
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function  procesaVariables2(){
 	//alert("entro");
   if (conexion2.readyState == 4){
	   if(conexion2.responseText == 1 || conexion2.responseText == 0 ){
		  	alert("No se encontraron registros!");
		}else{
		   	document.getElementById('gcontable').innerHTML = conexion2.responseText;
			
		}
	oculta();
  }
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function  procesaVariables3(){
 	//alert("entro");
   if (conexion2.readyState == 4){
	   if(conexion2.responseText == 1 || conexion2.responseText == 0 ){
		  	alert("No se encontraron registros!");
		}else{
		   	document.getElementById('sectorinst').innerHTML = conexion2.responseText;
			
		}
	oculta();
  }
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function ExcelSICON(valor,estadofinanciero){
	
	var x = document.getElementsByName("anios");
	var i;
	var strForText="";
	var SosText="";	


	var elem = document.getElementById("cboGruCont");
	if (elem === null) alert('Seleccione un Sector y a continuación un grupo.');
	else
	{
		var strAnio = document.getElementById("cboAnio").value;
		var strFor = document.getElementById("cboFormato").value;
		var strSec = document.getElementById("cboSecInt").value;
		var strGru = document.getElementById("cboGruCont").value;

	    var elt = document.getElementById("cboFormato");
	    var strForText= elt.options[elt.selectedIndex].text;
		
		var cbos = document.getElementById("cboSociedad");
		SosText= cbos.options[cbos.selectedIndex].text;

		var cgru = document.getElementById("cboGruCont");
		GruText= cgru.options[cgru.selectedIndex].text;

		if (strGru==""){
				alert("Seleccione al menos un grupo");
		}
		else
		{
			document.fmrEF.action='InfFinancieraactecon_ReporteExcel.asp'+"?anio="+strAnio+"&for="+strFor+"&sec="+strSec+"&ForText="+strForText+"&strGru="+strGru+"&SosText="+SosText+"&GruText="+GruText;
			document.fmrEF.submit();
			/*document.fmrEF.action="InfFinancieraactecon.asp";*/
			document.fmrEF.target='_self';

		}
	}
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaDividendos()
{
	var xannio=document.getElementById('cboAnio').value;

	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}

	conexion2=Ajax();
	muestra();
	var url= ''

	url='Dividendos_Reporte.asp';
	url=url+'?annio='+xannio;

	conexion2.open('POST', url, true);
	//alert(noCache(url));
	conexion2.setRequestHeader('Content-Type', 'text/html');
	conexion2.setRequestHeader('encoding', 'iso-8859-1');
	conexion2.onreadystatechange = procesaVariables;
	conexion2.send(null);	
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function cargaDividendosExcel()
{
	var xannio=document.getElementById('cboAnio').value;
	
	if (xannio=='')  {
		alert ('Seleccione Año');
		document.getElementById('cboAnio').focus();
		return false;
	}
	
	var url= 'Dividendos_ReporteExcel.asp';

	url=url+'?annio='+xannio;
	//alert(url);
	document.fmrEF.action=url;
	document.fmrEF.submit();
	document.fmrEF.target='_self';
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function CargaFiltro(cboName,url,FuncDepend)
{
	firstValue=0
	var c = document.getElementById(cboName);
	c.innerHTML = "";
	$.ajax({
	    url: url,
	    type: 'POST',
	    success: function(data) {
	    	for (var i = 0; i < data.length-1; i++) {
				var option = document.createElement("option");
				option.text = data[i].des;					
				option.value =  data[i].cod;
				c.add(option);
	    	}
			if (FuncDepend!="")
			{
    			call_others(FuncDepend);
			}
	    },
	    error: function() {
	    	alert("Ocurrió un error. Comuníquese con el administrador del sistema.");
	    },
	    cache: false,contentType: "json; charset:ISO-8859-1",processData: false
	}, 'json');
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
function call_others(function_name) {
	window[function_name]();
	//eval(function_name+"()");
}
/*-----------------------------------------------------------------------------------------------------------------------------------------------*/
