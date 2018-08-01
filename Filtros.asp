<!--#include file="Conexion.asp"-->
<%
	rep=Request.QueryString("rep")
	
	eeff=Request.QueryString("eeff")
	anio=Request.QueryString("anio")
	per=Request.QueryString("per")
	gcon=Request.QueryString("gcon")
	SecSic=Request.QueryString("secsic")
	niv=Request.QueryString("niv")
	data=Request.QueryString("data")
	
	SQL = ""

	Select Case rep
	Case "anio"
		SQL = "select distinct anio as cod, anio as des from DirEmpresasSICON order by anio desc"
	Case "anioBC"
		SQL = "select distinct anio as cod, anio as des from SICONPRO_BALCOM_1 order by anio desc"
	'-------------
	Case "anioAnexo"
		SQL = "SELECT DISTINCT ANO_EJE AS cod, ANO_EJE as des FROM "& data &" order by ANO_EJE DESC"
	Case "gcont_Anexo"
		SQL = "sp_lista_GrupoContable_Anexo "& data &","& anio	
	Case "codigo_Anexo"
		SQL = "sp_lista_codigos_Anexo "& anio &",'"& gcon &"',"& niv
	'----------------------		
	Case "gcont_RepAnio"
		SQL = "sp_lista_GrupoContable_Anioeeff "& anio &","& eeff
	Case "gcontBC_RepAnio"
		SQL = "sp_lista_GrupoContableBC_Anioeeff "& anio
	Case "SecSicon_RepAnio"
		SQL = "sp_lista_SectorSicon_Anioeeff "& anio &","& eeff
	Case "codigo_AnioeeffGConNiv"
		SQL = "sp_lista_codigos_AnioeeffGrupoNiv "& anio &",'"& eeff &"','"& gcon &"',"& niv
	Case "codigoBC_AnioGConNiv"
		SQL = "sp_lista_codigosBC_AnioGrupoNiv "& anio &",'"& gcon &"',"& niv
	Case "codigo_AnioeeffSecSicNiv"
		SQL = "sp_lista_codigos_AnioeeffSecSicNiv "& anio &",'"& eeff &"','"& SecSic &"',"& niv	

	End Select

	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open SQL , con

	Response.ContentType = "application/json; charset=ISO-8859-1"
	
	Response.Write("[")

	If Not rs.BOF Then rs.MoveFirst
	Do While Not rs.EOF

	Response.Write("{ ""cod"": """&rs("cod")&""", ""des"": """&rs("des")&"""},")
	rs.MoveNext
	Loop

	Response.Write("{ }")
	Response.Write("]")

	rs.Close
	Set rs = Nothing
	SQL=""
%>
