Attribute VB_Name = "ImpresiónPDF"
Sub Imprimir_PDF()

Dim Hoja As Object
Dim cont As Integer

Set Hoja = Worksheets(1)

Archivo = RutaPDF & "\" & NombrePOD & ".pdf"
'Archivo = RutaPDF & "/" & NombrePOD & ".pdf"

HojaOrigen.ExportAsFixedFormat Type:=xlTypePDF, Filename:=Archivo, Openafterpublish:=True


End Sub


