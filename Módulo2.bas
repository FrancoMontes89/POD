Attribute VB_Name = "Módulo2"
Option Explicit

Sub Actualizar_Lista()
Attribute Actualizar_Lista.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Actualizar_Lista Macro
'

'
    Range("H2").Select
    Selection.ListObject.QueryTable.Refresh BackgroundQuery:=False
End Sub
