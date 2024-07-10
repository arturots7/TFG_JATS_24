VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

'Si se produce un cambio en la selección del campo Alumnado del formulario Datos del alumnado
If Target.Address = "$C$8" Then
    'Se borra el contenido del campo Módulo solicitado del formuario Datos del alumnado
    Range("C16") = ""
    'Se selecciona la celda del campo Módulo solicitado del formulario Datos del alumnado
    Range("C16").Select
End If

'Si se produce un cambio en la selección del campo Módulo solicitado del formulario Datos del alumnado
If Target.Address = "$C$16" Then
    'Se borra el contenido del campo Título - Módulo/s aportado del formuario Datos del alumnado
    Range("C17").ClearContents
    'Se selecciona la celda del campo Título - Módulo/s aportado del formulario Datos del alumnado
    Range("C17").Select
End If

End Sub
