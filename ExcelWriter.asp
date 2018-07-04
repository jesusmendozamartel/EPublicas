<% 
Set appExcel = Server.CreateObject("Excel.Application")'Componente que se instala cuando posee office instalado sino NO SIRVE

'appExcel.Visible = false  

appExcel.Workbooks.Open("D:\Cesar\Fuentes\Fuente_Sicon\AplicativoSunat\FRM_ACTIVIDAD.xlsx")'abres el archivo de excel de tu maquina



appExcel.Sheets("AE_Modo1_BC_C").Select
For i = 30 To 135 ' empiezo a escribir desde la linea 30 hasta la 135
appExcel.Range("Y" & i).Value = "PEPE"
appExcel.Range("Z" & i).Value = "=4+5"
appExcel.Range("AA" & i).Value = "Resultados PDT"
appExcel.Range("AB" & i).Formula = "=4+5"
Next



appExcel.Sheets("AE_Modo1_PDT_C").Select
appExcel.Range("IA30").Select
appExcel.ActiveCell.Formula = "AAA"
'appExcel.Range("B" & 2).Value = 1 'escribe 1 en la linea b2 de excel 
appExcel.ActiveCell.Borders.Color = RGB(0, 0, 0)
appExcel.ActiveCell.Font.Name = "Arial" 
appExcel.ActiveCell.Font.Bold = True
appExcel.ActiveCell.Font.Size = 12
appExcel.ActiveCell.Font.Color = vbGreen
appExcel.ActiveCell.Interior.ColorIndex = 44
appExcel.ActiveCell.ColumnWidth = 27


'appExcel.ActiveWorkbook.SaveAs(Server.MapPath("F.xls"))
appExcel.ActiveWorkbook.SaveAs ("D:\Cesar\Fuentes\Fuente_Sicon\FRM_ACTIVIDAD_new.xlsx") 'salvo elarchivo como quieras y donde quieras
appExcel.Workbooks.Close ' cierro el objeto y listo
appExcel.Quit

Set appExcel = Nothing
Set objWorkbook = Nothing
Set objWorksheet = Nothing
Set colSheets = Nothing

%>