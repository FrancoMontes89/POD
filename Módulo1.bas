Attribute VB_Name = "M�dulo1"
Option Explicit

Public cnn As ADODB.Connection
Public recordset As ADODB.recordset


Sub Abrir_Conexi�n_ADODB()

Dim query As String
Dim database As String

Set cnn = New ADODB.Connection
Set recordset = New ADODB.recordset

query = "SELECT * FROM [Insumos];"
database = "https://esustentables.sharepoint.com/sites/OyM/Caada%20Honda/02-Documentaci�n general O&M/01-Operaci�n/01-POD/02-Hist�ricos/POD_Ca�ada Honda.accdb"

With cnn
.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://esustentables.sharepoint.com/sites/OyM/Caada%20Honda/02-Documentaci�n general O&M/01-Operaci�n/01-POD/02-Hist�ricos/POD_Ca�ada Honda.accdb;"
.Open
End With

Debug.Print cnn.State

recordset.Open query, cnn, adOpenDynamic, adLockOptimistic


End Sub

