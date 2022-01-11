Attribute VB_Name = "SQL"
Sub M01_resumen()

'Referencia Microsoft ActiveX Data Objets 6.0 Library

strFileName = ThisWorkbook.FullName


'Inicialización de variables
    Dim con As New ADODB.Connection
    Dim rs As New ADODB.Recordset
    Dim ConnectionString As String
    Dim sql As String

    'Cadena de conexión con una Base de Datos
'ConnectionStringMySQL = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" _
'& strFileName & ";Extended Properties=Excel 12.0 Xml;HDR=Yes;"

ConnectionStringMySQL = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" _
+ strFileName + ";Extended Properties=""Excel 12.0;IMEX=1;"""

    'Abrir conexión con la BBDD
    con.Open ConnectionStringMySQL
    
    'Timeout en segundos para ejecutar la SQL completa antes de reportar un error
    con.CommandTimeout = 900

    'Esta es la SQL que queremos consultar
    strsql = "TRANSFORM Sum([DATOS$].[VIAJES]) AS SumaDeContador SELECT [DATOS$].[CHAPA], Sum([DATOS$].[VIAJES]) AS [Totales por mes] FROM [DATOS$] GROUP BY [DATOS$].[CHAPA] PIVOT [DATOS$].[TIPO DE MATERIAL]"
     
    'Ejecutamos la consulta
     rs.Open strsql, con
    
    'Copiamos los resultados de la SQL sobre la hoja del Excel en la celda A2
   
Sheets("Resumen").Range("A3").CopyFromRecordset rs

'ponemos los nombres de los campos del recordset
For n = 0 To (rs.Fields.Count - 1)
Sheets("Resumen").Cells(2, n + 1) = rs.Fields(n).Name
Next
    
    'Cerramos las conexiones
    rs.Close
    con.Close
    
End Sub


