<% 
Set appExcel = Server.CreateObject("Excel.Application")'Componente que se instala cuando posee office instalado sino NO SIRVE

appExcel.Workbooks.Open("D:\Cesar\Fuentes\Fuente_Sicon\AplicativoSunat\Libro2.xlsx")'abres el archivo de excel de tu maquina
appExcel.Range("B" & 2).Value = 1 'escribe 1 en la linea b2 de excel 
For i = 4 To 10 ' empiezo a escribir desde la linea 4 hasta la 10 
appExcel.Range("A" & i).Value = "PEPE"
appExcel.Range("B" & i).Value = 15
appExcel.Range("C" & i).Value = "Profesional"
appExcel.Range("D" & i).Value = 40
Next
appExcel.ActiveWorkbook.SaveAs ("D:\Cesar\Fuentes\Fuente_Sicon\FRM_ACTIVIDAD_new.xlsx") 'salvo elarchivo como quieras y donde quieras
appExcel.Workbooks.Close ' cierro el objeto y listo

%>