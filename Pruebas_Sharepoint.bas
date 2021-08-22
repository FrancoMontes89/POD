Attribute VB_Name = "Pruebas_Sharepoint"
Option Explicit
Public conn As ADODB.Connection
Public recordset As ADODB.recordset

Sub CopiarRegistrosList()

Dim conn As ADODB.Connection
Dim rst As ADODB.recordset
Dim query As String
Dim Fecha As Date
Dim Edad As Integer

Fecha = "15 / 7 / 1992"
Edad = 29


Set conn = New ADODB.Connection
Set rst = New ADODB.recordset

query = "SELECT * FROM [ListaPrueba];"

With conn

    .ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://esustentables.sharepoint.com/sites/OyM;LIST={690d4623-1078-473b-8d21-972820ada2b6};"
    .Open

End With

rst.Open query, conn, adOpenDynamic, adLockOptimistic

    rst.AddNew
        rst!ASD = Fecha
        rst!CVB = Edad
    rst.Update ' commit changes to SP list

    If CBool(rst.State And adStateOpen) = True Then rst.Close
    If CBool(conn.State And adStateOpen) = True Then conn.Close
End Sub




