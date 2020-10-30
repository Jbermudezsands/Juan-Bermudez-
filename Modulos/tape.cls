VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTape"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************
' Módulo de clase de la aplicación de ejemplo VCR
' Actúa como mecanismo de transporte de cinta del
' VCR y controla la reproducción de las imágenes
'******************************************************

Option Explicit

Public Left As Integer      'Última posición
Public Forward As Boolean   'Dirección de la cinta
Public Speed As Integer     'Velocidad de la cinta
'******************************************************
' Propósito:  Calcula las nuevas coordenadas
'           de cada paso de la animación.
' Entradas:   Width: la anchura del cuadro de imagen
'           en el que se presenta la animación
'******************************************************
Public Sub Animate(Width As Integer)
    If Forward = True Then
        ' hacia delante - incrementa la posición izquierda actual
        ' hasta llegar al borde derecho
        If Left < Width - 50 Then
            Left = Left + 50
        Else
            Left = 0
        End If
    Else
        ' hacia delante - decrementa la posición izquierda actual
        ' hasta llegar al borde izquierdo
        If Left > 0 Then
            Left = Left - 50
        Else
            Left = Width - 50
        End If
    End If
End Sub

Private Sub Class_Initialize()
    ' Inicializa las propiedades de la clase
    Forward = True
    Left = 0
    Speed = 300
End Sub
