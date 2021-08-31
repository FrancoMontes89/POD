Attribute VB_Name = "ADO"
Option Explicit

Sub AbrirConexiónLista() '(ByVal NombreLista As String, GUID As String)

Set ConnADODB = New ADODB.Connection
Set rst = New ADODB.recordset

query = "SELECT * FROM [" & NombreLista & "];"

With ConnADODB
       .ConnectionString = _
       "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://esustentables.sharepoint.com/sites/OyM;LIST={" & GUID & "};"
       .Open
End With
    
rst.Open query, ConnADODB, adOpenDynamic, adLockOptimistic

End Sub

