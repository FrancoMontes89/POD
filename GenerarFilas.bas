Attribute VB_Name = "GenerarFilas"

Dim cont1 As Integer
Dim ConTareaMant As Integer


Sub T1_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False

Tabla = 1
ProximaTabla = Tabla + 1
'fila = Application.Match(Tabla, Worksheets("RE-OPE-03_Rev.01").Columns(1), 0)
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0)
'fila_proximatabla = Application.Match(ProximaTabla, Worksheets("RE-OPE-03_Rev.01").Columns(1), 0) - 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1


Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila + 1, 1), Cells(fila + 1, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial

End Sub


Sub T3_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False

Tabla = 3
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0)
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila + 3, 1), Cells(fila + 3, 9)).Copy
Range(Cells(fila_proximatabla - 3, 1), Cells(fila_proximatabla - 3, 9)).PasteSpecial

End Sub


Sub T4_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False

'Resta definir como automatizarlo =4
Tabla = 4
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial

End Sub

Sub T5_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False

'Resta definir como automatizarlo =4
Tabla = 5
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial

End Sub

Sub T6_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False

'Resta definir como automatizarlo =4
Tabla = 6
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 2
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial

End Sub

Sub T7_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String

Application.ScreenUpdating = False

Tabla = 7
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial
ConTareaMant = ConTareaMant + 1


End Sub

Sub T8_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String
Dim ConConsumible As Integer

Application.ScreenUpdating = False

Tabla = 8
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial
ConConsumible = ConConsumible + 1


End Sub

Sub T9_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False
'Resta definir como automatizarlo =4
Tabla = 9
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial

End Sub

Sub T10_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False
'Resta definir como automatizarlo =4
Tabla = 10
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0) - 1

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial

End Sub

Sub T11_Sumar()

Dim Tabla As Integer
Dim ProximaTabla As Integer
Dim fila As String
Dim fila_proximatabla As String


Application.ScreenUpdating = False

'Resta definir como automatizarlo =4
Tabla = 11
ProximaTabla = Tabla + 1
fila = Application.Match(Tabla, Worksheets("Parte operativo diario").Columns(1), 0) + 1
fila_proximatabla = Application.Match(ProximaTabla, Worksheets("Parte operativo diario").Columns(1), 0)

Range(Cells(fila_proximatabla - 1, 1), Cells(fila_proximatabla - 1, 1)).EntireRow.Insert
fila_proximatabla = fila_proximatabla + 1
Range(Cells(fila, 1), Cells(fila, 9)).Copy
Range(Cells(fila_proximatabla - 2, 1), Cells(fila_proximatabla - 2, 9)).PasteSpecial

End Sub


