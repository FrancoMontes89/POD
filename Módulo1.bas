Attribute VB_Name = "Módulo1"
Option Explicit

Public cnn As ADODB.Connection
Public recordset As ADODB.recordset


Sub Abrir_Conexión_ADODB()

Dim query As String
Dim database As String

Set cnn = New ADODB.Connection
Set recordset = New ADODB.recordset

query = "SELECT * FROM [Insumos];"
database = "https://esustentables.sharepoint.com/sites/OyM/Caada%20Honda/02-Documentación general O&M/01-Operación/01-POD/02-Históricos/POD_Cañada Honda.accdb"

With cnn
.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://esustentables.sharepoint.com/sites/OyM/Caada%20Honda/02-Documentación general O&M/01-Operación/01-POD/02-Históricos/POD_Cañada Honda.accdb;"
.Open
End With

Debug.Print cnn.State

recordset.Open query, cnn, adOpenDynamic, adLockOptimistic


End Sub

