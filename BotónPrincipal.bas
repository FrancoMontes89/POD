Attribute VB_Name = "BotónPrincipal"
Option Explicit

Public FechaPOD As Date
Public UserObject, IdPCObject As Object
Public user, IdPC, Ruta, PS, Archivo, NombrePOD As String
Public POD, BDD As Workbook
Public HojaOrigen, HojaDestino As Excel.Worksheet
Public RangoOrigen, RangoDestino, Área_Impresión As Excel.Range
Public ufila, ucolumna, dd, aa As Integer
Public ntabla, filatabla, utabla, nfilas, fila_proximatabla, preal_campo1, p_proximotítulo As Integer
Public cont As Integer
Public p_titulo, p_enunciado, p_campo1, ncampos, RutaPDF As Variant


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
RutaPDF = Application.WorksheetFunction.VLookup(PS, Sheets(2).Range("A1:G6"), 5, False)
        'Mensaje de fecha faltante
        If FechaPOD = 0 Then
        MsgBox "No ha indicado una fecha para el POD, por favor complete el campo fecha y continue.", , "Fecha no definida"
        Exit Sub
        End If

'Calcula cuantas "tablas" tiene el POD por defecto
utabla = Cells(ActiveSheet.Cells(Rows.Count, "A").End(xlUp).Row, 1).Value

'Define y abre el archivo que oficiará como Base de datos

Ruta = Application.WorksheetFunction.VLookup(PS, Sheets(2).Range("A1:G6"), 4, False)

Set BDD = Workbooks.Open(Ruta)
ActiveWindow.WindowState = xlMinimized

Set HojaDestino = BDD.Sheets(1)

HojaDestino.Unprotect Password:="360O&M2020"

ufila = BDD.Sheets(1).Cells(Rows.Count, "B").End(xlUp).Row + 1

        'Mensaje de advertencia de sobre escritura:
        If HojaDestino.Cells(ufila - 1, 1).Value = FechaPOD Then
        MsgBox "El último registro de la base de datos coincide con la fecha que se desea ingresar," & vbNewLine & "Por favor corrobore que la fecha del POD actual sea la correcta.", , "Advertencia de sobre escritura"
        Exit Sub
        Else
        End If

Dim Hoja As Object
Dim cont As Integer


'Área de impresión
Set Área_Impresión = Range(Cells(1, 1), Cells(ufila + 2, 10))
ActiveSheet.PageSetup.PrintArea = Área_Impresión.Address

'Control ortográfico
If Área_Impresión.CheckSpelling = False Then
Exit Sub
End If

'Set Hoja = Worksheets(1)
Archivo = RutaPDF & "\" & NombrePOD & ".pdf"

HojaOrigen.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Archivo, Openafterpublish:=True
'End Sub

If MsgBox("Está a punto de compartir el POD a los interesados." & vbNewLine & "Por favor, revise el PDF generado, ciérrelo y confirme el envío haciendo click en el botón 'SI'" & vbNewLine & "EN CASO CONTRARIO: Recuerde eliminar el PDF creado en el sitio de Sharepoint antes de reintentar!", vbYesNo, "Confirmación de envío") = 6 Then
    Call Enviar_mail
    Else
    ThisWorkbook.FollowHyperlink RutaPDF
    End
End If

'Copiado:

ntabla = 1

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 7).Value
        .Cells(ufila + cont, 6).Value = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

ntabla = 2

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"


'Copia tabla4

ntabla = 4

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        .Cells(ufila + cont, 6).Value = HojaOrigen.Cells(preal_campo1 + cont, 7).Value
        .Cells(ufila + cont, 7).Value = HojaOrigen.Cells(preal_campo1 + cont, 8).Value
        .Cells(ufila + cont, 8).Value = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia tabla5

ntabla = 5

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"


filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia tabla6

ntabla = 6

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_titulo = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 3, False)
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        .Cells(ufila + cont, 6).Value = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        .Cells(ufila + cont, 7).Value = HojaOrigen.Cells(preal_campo1 + cont, 7).Value
        .Cells(ufila + cont, 8).Value = HojaOrigen.Cells(preal_campo1 + cont, 8).Value
        .Cells(ufila + cont, 9).Value = HojaOrigen.Cells(preal_campo1 + cont, 9).Value

        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia tabla7
ntabla = 7

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        .Cells(ufila + cont, 6).Value = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        .Cells(ufila + cont, 7).Value = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia tabla8
ntabla = 8

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia tabla9
ntabla = 9

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 9).Value
        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia tabla10

ntabla = 10

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_titulo = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 3, False)
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1

cont = 0

    Do While cont < ncampos + 1
      If (HojaOrigen.Cells(preal_campo1 + cont, 3).Value <> "N/A") Or Not IsEmpty(HojaOrigen.Cells(preal_campo1 + cont, 7).Value) Then
        With HojaDestino
          .Cells(ufila, 1).Value = FechaPOD
          .Cells(ufila, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
          .Cells(ufila, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
          .Cells(ufila, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
          .Cells(ufila, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
          .Cells(ufila, 6).Value = HojaOrigen.Cells(preal_campo1 + cont, 7).Value
        End With
      ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
      cont = cont + 1
      Else
      cont = cont + 1
      End If
    cont = cont + 1
    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia tabla11
ntabla = 11

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"

filatabla = HojaOrigen.Range("A:A").Find(ntabla, LookIn:=xlValues, LookAt:=xlWhole).Row
fila_proximatabla = HojaOrigen.Range("A:A").Find(ntabla + 1, LookIn:=xlValues, LookAt:=xlWhole).Row
p_proximotítulo = Application.WorksheetFunction.VLookup(ntabla + 1, Sheets(2).Range("O1:T13"), 3, False)
p_enunciado = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 4, False)
p_campo1 = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 5, False)
ncampos = Application.WorksheetFunction.VLookup(ntabla, Sheets(2).Range("O1:T13"), 6, False)

ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1

preal_campo1 = filatabla + p_campo1
ncampos = fila_proximatabla + p_proximotítulo - preal_campo1

cont = 0

    Do While cont < ncampos

      With HojaDestino
        .Cells(ufila + cont, 1).Value = FechaPOD
        .Cells(ufila + cont, 2).Value = HojaOrigen.Cells(preal_campo1 + cont, 2).Value
        .Cells(ufila + cont, 3).Value = HojaOrigen.Cells(preal_campo1 + cont, 3).Value
        .Cells(ufila + cont, 4).Value = HojaOrigen.Cells(preal_campo1 + cont, 4).Value
        .Cells(ufila + cont, 5).Value = HojaOrigen.Cells(preal_campo1 + cont, 5).Value
        .Cells(ufila + cont, 6).Value = HojaOrigen.Cells(preal_campo1 + cont, 6).Value
        .Cells(ufila + cont, 7).Value = HojaOrigen.Cells(preal_campo1 + cont, 9).Value

        End With

    cont = cont + 1

    Loop

'Borra la fila excedente
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
HojaDestino.Rows(ufila).EntireRow.Delete
HojaDestino.Protect Password:="360O&M2020"

'Copia registro escritura
ntabla = 12

Set HojaDestino = BDD.Sheets(ntabla)
HojaDestino.Unprotect Password:="360O&M2020"
ufila = HojaDestino.Cells(Rows.Count, "B").End(xlUp).Row + 1
        With HojaDestino
        .Cells(ufila, 1).Value = FechaPOD
        .Cells(ufila, 2).Value = IdPC
        .Cells(ufila, 3).Value = user
        End With

HojaDestino.Protect Password:="360O&M2020"
BDD.Save
BDD.Close Savechanges:=True

            'Mensaje de confirmación de cierre
            If MsgBox("El POD se envió correctamente," & vbNewLine & "Si no desea crear un nuevo POD haga click en el botón Aceptar, caso contrario seleccione Cancelar", vbOKCancel, "Confirmación de envío") = 1 Then
            ActiveWorkbook.Close Savechanges:=False

            Else
            End
            End If


End Sub
