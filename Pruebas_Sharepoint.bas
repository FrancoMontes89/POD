Attribute VB_Name = "Pruebas_Sharepoint"
Option Explicit
Public conn As ADODB.Connection
Public recordset As ADODB.recordset

Sub AbrirDocSharepointDesktopApp()

Dim DirecciónSharepoint As String

DirecciónSharepoint = "https://esustentables.sharepoint.com/:x:/s/OyM/EYXkjTgo59pAsw0ohquPyX0BJ9VXJj2jaYlS65jRAv5fFA?e=lr1Pdz"

ActiveWorkbook.FollowHyperlink Address:=DirecciónSharepoint

End Sub

Sub Conexión_Access_Sharepoint()

'On Error GoTo Errores
Dim host, database, user, pass, query, strcon As String
'host = ""
'database = ""
'user = ""
'pass = ""
Set conn = New ADODB.Connection
'conn.Open "Driver={Microsoft Access Driver (*.mdb)}; Server= & host & ";Database=" & database & ";Uid=" & user & ";Pwd=" & pass & ";""
'Debug.Print "La conexión se ha realizado correctamente"
'Exit Sub

strcon = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=https://esustentables.sharepoint.com/sites/OyM/Caada%20Honda/02-Documentación general O&M/01-Operación/01-POD/02-Históricos/POD_Cañada Honda.accdb;"
conn.Open (strcon)

query = "SELECT * FROM Table1"
recordset.Open qry, conn, adOpenKeyset

recordset.Close
conn.Close

'Errores:
'
'Msgbox Err Description, vbCritical


End Sub
