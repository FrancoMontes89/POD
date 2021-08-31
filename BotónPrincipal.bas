Attribute VB_Name = "BotónPrincipal"
Option Explicit

Public FechaPOD As Date
Public UserObject, IdPCObject As Object
Public user, IdPC, Ruta, PS, Archivo, NombrePOD, Indice, Aviso As String
Public POD, BDD As Workbook
Public HojaOrigen, HojaDestino As Excel.Worksheet
Public RangoOrigen, RangoDestino, Área_Impresión As Excel.Range
Public ufila, ucolumna, dd, aa As Integer
Public ntabla, filatabla, utabla, nfilas, fila_proximatabla, preal_campo1, p_proximotítulo As Integer
Public cont As Integer
Public p_titulo, p_enunciado, p_campo1, ncampos, RutaPDF As Variant
'Variables para el copiado a Listas:
Public ConnADODB As ADODB.Connection
Public rst As ADODB.recordset
Public NombreLista, GUID, query As String

Sub Copiar()

Application.ScreenUpdating = False

'Variables para definir los rangos a copiar
Set POD = ThisWorkbook
Set HojaOrigen = POD.Worksheets(1)

'Guarda nombre de usuario, Id PC y fecha POD
Set UserObject = CreateObject("Wscript.network")
user = UserObject.UserName
Set IdPCObject = CreateObject("Wscript.network")
IdPC = IdPCObject.ComputerName
FechaPOD = HojaOrigen.Range("I4").Value
PS = HojaOrigen.Range("B3").Value
NombrePOD = Year(FechaPOD) & "." & Month(FechaPOD) & "." & Day(FechaPOD) & " - " & PS
RutaPDF = Application.WorksheetFunction.VLookup(PS, Worksheets("REF").Range("A1:G6"), 5, False)
        'Mensaje de fecha faltante
        If FechaPOD = 0 Then
        MsgBox "No ha indicado una fecha para el POD, por favor complete el campo fecha y continue.", , "Fecha no definida"
        Exit Sub
        End If

'Calcula cuantas "tablas" tiene el POD por defecto
utabla = Cells(ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row, 1).Value '¿Qué utilidad tiene?
ufila = ActiveSheet.Cells(Rows.Count, "B").End(xlUp).Row

Dim Hoja As Object
Dim cont As Integer

'Área de impresión:
Set Área_Impresión = Range(Cells(1, 1), Cells(ufila + 2, 10))
ActiveSheet.PageSetup.PrintArea = Área_Impresión.Address

Control ortográfico:
If Área_Impresión.CheckSpelling = False Then
Exit Sub
End If

Set Hoja = Worksheets(1)
Archivo = RutaPDF & "\" & NombrePOD & ".pdf"

'HojaOrigen.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Archivo, Openafterpublish:=True
'End Sub


If MsgBox("Está a punto de compartir el POD a los interesados." & vbNewLine & "Por favor, revise el PDF generado, ciérrelo y confirme el envío haciendo click en el botón 'SI'" & vbNewLine & "EN CASO CONTRARIO: Recuerde eliminar el PDF creado en el sitio de Sharepoint antes de reintentar!", vbYesNo, "Confirmación de envío") = 6 Then
    'Call Enviar_mail
    Else
    ThisWorkbook.FollowHyperlink RutaPDF
    End
End If

'Copiado:

ntabla = 1

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

    Do While cont < ncampos

        rst.AddNew
            rst!Rol = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
            rst!Modo = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
            rst!nombre = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
            rst!Inicio = Int(HojaOrigen.Cells(preal_campo1 + cont, 7).Value * 1440)
            rst!Fin = Int(HojaOrigen.Cells(preal_campo1 + cont, 9).Value * 1440)
            rst!Fecha = FechaPOD
        rst.Update

    cont = cont + 1

    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

ntabla = 2

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1


Aviso = HojaOrigen.Cells(preal_campo1 + cont, 2).Value

If Aviso = "" Then
GoTo 4
Else

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos

        rst.AddNew
            rst!Fecha = FechaPOD
            rst!Aviso = Aviso
        rst.Update
        cont = cont + 1
    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

End If

4:

'Copia tabla4

ntabla = 4

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos
        rst.AddNew
            rst.Fields("Fecha") = FechaPOD
            rst.Fields("Parque") = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
            rst.Fields("Generada Medidor Principal") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
            rst.Fields("Generada Medidor Control") = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
            rst.Fields("Generada Analizador de redes") = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
            rst.Fields("Consumida Medidor Principal") = HojaOrigen.Cells(preal_campo1 + cont, 7).Value
            rst.Fields("Consumida Medidor Control") = HojaOrigen.Cells(preal_campo1 + cont, 8).Value
            rst.Fields("Consumida Analizador de redes") = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        rst.Update
        cont = cont + 1
    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

'Copia tabla5

ntabla = 5

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos

        rst.AddNew
        rst.Fields("Fecha") = FechaPOD
        rst.Fields("Hora") = Int(HojaOrigen.Cells(preal_campo1 + cont, 2).Value * 1440)
        rst.Fields("Código") = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        rst.Fields("Parque") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        rst.Fields("Novedad") = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        rst.Update

    cont = cont + 1

    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

'Copia tabla6

ntabla = 6

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos

        rst.AddNew

        rst.Fields("Parque") = PS
        rst.Fields("Fecha") = FechaPOD
        rst.Fields("Componente") = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        rst.Fields("Tipo(Int/Ext)") = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        rst.Fields("Inicio en jornada") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        rst.Fields("Hora inicio") = Int(HojaOrigen.Cells(preal_campo1 + cont, 5).Value * 1440)
        rst.Fields("Hora Fin") = Int(HojaOrigen.Cells(preal_campo1 + cont, 6).Value * 1440)
        rst.Fields("Potencia Comprometida") = HojaOrigen.Cells(preal_campo1 + cont, 7).Value
        rst.Fields("Duración estimada") = Int(HojaOrigen.Cells(preal_campo1 + cont, 8).Value * 1440)
        rst.Fields("No Suministrada") = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        rst.Update

        cont = cont + 1

    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

'Copia tabla7
ntabla = 7

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos

        rst.AddNew
        rst.Fields("Fecha") = FechaPOD
        rst.Fields("Parque") = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        rst.Fields("Componente") = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        rst.Fields("Tarea") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        rst.Fields("Status") = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        rst.Fields("Comentarios") = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        rst.Fields("Tipo tarea") = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        rst.Update

    cont = cont + 1

    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

'Copia tabla8
ntabla = 8

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos

        rst.AddNew
        rst.Fields("Parque") = PS
        rst.Fields("Fecha") = FechaPOD
        rst.Fields("Cantidad consumida") = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        rst.Fields("Código SAP") = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        rst.Fields("Descripción") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        rst.Fields("Tarea asociada") = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        rst.Update

    cont = cont + 1

    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

'Copia tabla9

ntabla = 9

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos

        rst.AddNew
        rst.Fields("Parque") = PS
        rst.Fields("Fecha") = FechaPOD
        rst.Fields("Sistema") = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        rst.Fields("Estado") = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        rst.Fields("Observaciones") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        rst.Fields("Reclamo") = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        rst.Update

    cont = cont + 1

    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

'Copia tabla10

ntabla = 10

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos
    rst.AddNew
    rst.Fields("Parque") = PS
    rst.Fields("Fecha") = FechaPOD
    rst.Fields("Categoría") = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
    rst.Fields("Referencia") = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
    rst.Fields("Número") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
    rst.Fields("Hora") = Int(HojaOrigen.Cells(preal_campo1 + cont, 6).Value * 1440)
    rst.Fields("Detalle") = HojaOrigen.Cells(preal_campo1 + cont, 7).Value
    rst.Update
    cont = cont + 1
    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close

'Copia tabla11
ntabla = 11

Indice = PS & "_" & ntabla
filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Worksheets("REF").Range("Q1:V13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Worksheets("REF").Range("Q1:V13"), 5, False)
NombreLista = Application.WorksheetFunction.VLookup(Indice, [Listas], 2, False)
GUID = Application.WorksheetFunction.VLookup(Indice, [Listas], 3, False)

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1 - 1

Call ADO.AbrirConexiónLista

cont = 0

    Do While cont < ncampos

        rst.AddNew
        rst.Fields("Parque") = PS
        rst.Fields("Fecha") = FechaPOD
        rst.Fields("Hora Ingreso") = Int(HojaOrigen.Cells(preal_campo1 + cont, 2).Value * 1440)
        rst.Fields("Empresa") = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        rst.Fields("Apellido") = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        rst.Fields("Nombre") = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        rst.Fields("Motivo") = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        rst.Fields("Hora Egreso") = Int(HojaOrigen.Cells(preal_campo1 + cont, 9).Value * 1440)
        rst.Update
    cont = cont + 1

    Loop

If CBool(rst.State And adStateOpen) = True Then rst.Close
If CBool(ConnADODB.State And adStateOpen) = True Then ConnADODB.Close


            'Mensaje de confirmación de cierre
            If MsgBox("El POD se envió correctamente," & vbNewLine & "Si no desea crear un nuevo POD haga click en el botón Aceptar, caso contrario seleccione Cancelar", vbOKCancel, "Confirmación de envío") = 1 Then
            ActiveWorkbook.Close Savechanges:=False

            Else
            End
            End If


End Sub
