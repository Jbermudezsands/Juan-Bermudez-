VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSubJustificacion 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepSubJustificacion.dsx":0000
End
Attribute VB_Name = "ArepSubJustificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalHoraEmpleado As String, TotalHorasDepartamento As String

Private Sub Detail_Format()
Dim TotalHoras As String

If Me.FldBegin.Text <> "" Then
    TotalHoras = DateDiff("s", Me.FldBegin.Text, Me.FldEnd.Text)
    Me.LblTotalHoras.Caption = Int(TotalHoras / 3600) & ":" & Int((TotalHoras Mod 3600) / 60)
    
    TotalHoraEmpleado = sumaHoras(Me.LblTotalHoras.Caption, TotalHoraEmpleado)
    TotalHorasDepartamento = sumaHoras(Me.LblTotalHoras.Caption, TotalHorasDepartamento)
    
    Me.LblFechaHora.Caption = Format(Me.FldBegin.Text, "dd/mm/yyyy") & "-" & Format(Me.FldEnd.Text, "dd/mm/yyyy")
End If

End Sub


