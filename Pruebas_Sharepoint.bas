Attribute VB_Name = "Pruebas_Sharepoint"
Option Explicit
Public Conn As ADODB.Connection
Public recordset As ADODB.recordset

Sub CopiarRegistrosList()

Dim Conn As ADODB.Connection
Dim rst As ADODB.recordset
Dim query As String
Dim Fecha As Date
Dim Edad As Integer

Set Conn = New ADODB.Connection
Set rst = New ADODB.recordset

'query = "SELECT * FROM [ListaPrueba];"
query = "SELECT * FROM [FIAMBALÁ_Personal];"

'49e29747-ee21-4f4e-a80c-0a1ca50b72b7
''690d4623-1078-473b-8d21-972820ada2b6


With Conn

    .ConnectionString = _
    "Provider=Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://esustentables.sharepoint.com/sites/OyM;LIST={49e29747-ee21-4f4e-a80c-0a1ca50b72b7};"
    .Open

End With

rst.Open query, Conn, adOpenDynamic, adLockOptimistic

rst.AddNew
'    rst!asd = Fecha
'    rst!CVB = Edad
            rst!Rol = "1" 'HojaOrigen.Cells(preal_campo1 + cont, 2).Value
            rst!Modo = "12" 'HojaOrigen.Cells(preal_campo1 + cont, 3).Value
            rst!nombre = "13" 'HojaOrigen.Cells(preal_campo1 + cont, 5).Value
            rst!Inicio = "14"  'HojaOrigen.Cells(preal_campo1 + cont, 7).Value
            rst!Fin = "15" 'HojaOrigen.Cells(preal_campo1 + cont, 9).Value
'            rst!Fecha = "12/12/2021"
'            rst!Rol = "1" 'HojaOrigen.Cells(preal_campo1 + cont, 2).Value
'            rst!Modo = "12" 'HojaOrigen.Cells(preal_campo1 + cont, 3).Value
'            rst!Nombre = "13" 'HojaOrigen.Cells(preal_campo1 + cont, 5).Value
'            rst!Inicio = "14"  'HojaOrigen.Cells(preal_campo1 + cont, 7).Value
'            rst!Fin = "15" 'HojaOrigen.Cells(preal_campo1 + cont, 9).Value
'            rst!Fecha = FechaPOD
rst.Update ' commit changes to SP list

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(Conn.State And adStateOpen) = True Then Conn.Close
    
End Sub




