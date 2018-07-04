<!--#include file="Conexion.asp"-->
<%
dim Tabla(5000,5000)
dim Tabla1(5000,5000)

Response.Charset= "ISO-8859-1" 
	
	anio=Request.QueryString("anio")
	formato=Request.QueryString("for")
	grupo=Request.QueryString("strGru")
	sector=Request.QueryString("sec")	
	ForText=Request.QueryString("ForText")	

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

	response.write("<table width='50%' border='0' cellspacing='0' cellpadding='0'><tr><td width='24%' valign='top'><table  class='tabla1' width='100%' border='0'>")



	Set objFieldsC = rsC.Fields
	response.write("<tr bgcolor='#FFE2C6'><td rowspan=4><strong>"&ForText&"</strong></td>")

	For intLoop = 0 To (objFieldsC.Count - 1)
		rsC.MoveFirst
		if intLoop > 0 Then
			response.write("<tr bgcolor='#FFE2C6'>")
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
		    response.write("<td colspan=2>"&rsC(intLoop)&"</td>")        	
		    rsC.MoveNext
		wend
		response.write("</tr>")
    Next




	Set objFieldsR = rsR.Fields
	Set objFieldsD = rsD.Fields

		response.write("<tr bgcolor='#FFE2C6'>")
	For intLoop = 0 To (objFieldsR.Count - 1)
        response.write("<td><strong>"&objFieldsR.Item(intLoop).Name&"</strong></td>")        	
    Next

	rsC.MoveFirst
    while not rsC.eof
	    For intLoop = 0 To (objFieldsD.Count - 1)
	    	response.write("<td><strong>"&objFieldsD.Item(intLoop).Name&"</strong></td>")        	
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
	        	response.write("<td align='left' style='mso-number-format:0.00;'>"&rsD(intLoop)&"</td>")
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
%>
