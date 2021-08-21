Attribute VB_Name = "EnviarMail"
Option Explicit

Sub Enviar_mail()

Dim AppOutlook As Outlook.Application
Dim Mail As Outlook.MailItem

Set AppOutlook = CreateObject("Outlook.Application")
Set Mail = AppOutlook.CreateItem(olMailItem)

On Error Resume Next
With Mail

.To = Application.WorksheetFunction.VLookup(PS, Sheets(2).Range("A1:M6"), 6, False)
.CC = Application.WorksheetFunction.VLookup(PS, Sheets(2).Range("A1:M6"), 7, False)
.BCC = ""
.Subject = "Parte operativo diario -" & PS & " d�a " & FechaPOD
.Body = "Estimados," & vbCrLf & "Por el presente se les adjunta el parte operativo diario del d�a " & FechaPOD & "." & vbCrLf & "La informaci�n que este contiene ha sido guardada en la base de datos correspondiente." & vbCrLf & "Cualquier consulta a disposici�n."
.attachments.Add Archivo
.Send

End With
End Sub
