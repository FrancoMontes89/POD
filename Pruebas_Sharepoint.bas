Attribute VB_Name = "Pruebas_Sharepoint"
Option Explicit
Public conn As ADODB.Connection
Public recordset As ADODB.recordset

Sub AbrirDocSharepointDesktopApp()

Dim Direcci�nSharepoint As String

Direcci�nSharepoint = "https://esustentables.sharepoint.com/:x:/s/OyM/EYXkjTgo59pAsw0ohquPyX0BJ9VXJj2jaYlS65jRAv5fFA?e=lr1Pdz"

ActiveWorkbook.FollowHyperlink Address:=Direcci�nSharepoint

End Sub

Sub Conexi�n_Access_Sharepoint()

'On Error GoTo Errores
Dim host, database, user, pass, query, strcon As String
'host = ""
'database = ""
'user = ""
'pass = ""
Set conn = New ADODB.Connection
'conn.Open "Driver={Microsoft Access Driver (*.mdb)}; Server= & host & ";Database=" & database & ";Uid=" & user & ";Pwd=" & pass & ";""
'Debug.Print "La conexi�n se ha realizado correctamente"
'Exit Sub

strcon = "Provider=Microsoft.ACE.OLEDB.12.0;" & "Data Source=https://esustentables.sharepoint.com/sites/OyM/Caada%20Honda/02-Documentaci�n general O&M/01-Operaci�n/01-POD/02-Hist�ricos/POD_Ca�ada Honda.accdb;"
conn.Open (strcon)

query = "SELECT * FROM Table1"
recordset.Open qry, conn, adOpenKeyset

recordset.Close
conn.Close

'Errores:
'
'Msgbox Err Description, vbCritical


End Sub
