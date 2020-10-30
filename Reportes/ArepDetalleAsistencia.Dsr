VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepDetalleAsistencia 
   Caption         =   "Detalle de Asistencias"
   ClientHeight    =   11010
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20340
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepDetalleAsistencia.dsx":0000
End
Attribute VB_Name = "ArepDetalleAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalLaboradas As String, TotalExtras As String, TotalEntrada As String, TotalSalida As String, TotalInicioComida As String, TotalFinComida As String

Private Sub ActiveReport_ReportStart()
 MDIPrimero.DtaEmpresa.Refresh
 Me.LblEmpresa.Caption = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
 Me.LblEmpresa1.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion")
 Me.LblEmpresa2.Caption = "RUC: " & MDIPrimero.DtaEmpresa.Recordset("NumeroRuc")
 RutaLogo = ""
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("RutaLogo")) Then
   RutaLogo = MDIPrimero.DtaEmpresa.Recordset("RutaLogo")
 End If
 Me.LblFechaImpreso.Caption = Format(Now, "DD/MM/YYYY")
 
 If (Dir(RutaLogo) <> "") Then
    Me.Logo.Picture = LoadPicture(RutaLogo)
 End If
 
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("MembreteLogo")) Then
   If MDIPrimero.DtaEmpresa.Recordset("MembreteLogo") = True Then
      Me.Logo.Width = Me.LblEmpresa.Width
      Me.Logo.Height = 700
      Me.PageSettings.TopMargin = 100
      Me.LblEmpresa.Top = 1000
      Me.LblEmpresa1.Top = 1300
      Me.LblEmpresa2.Top = 1550
      Me.Label15.Top = 1800
   End If
 End If
 
End Sub

Private Sub Detail_Format()

  If TotalLaboradas = "" Then
    TotalLaboradas = Me.FldLaboradas.Text
  Else
    TotalLaboradas = sumaHoras(TotalLaboradas, Me.FldLaboradas.Text)
 End If
 
 If TotalExtras = "" Then
    TotalExtras = Me.FldExtras.Text
  Else
    TotalExtras = sumaHoras(TotalExtras, Me.FldExtras.Text)
 End If
 
 If TotalEntrada = "" Then
    TotalEntrada = Me.FldEntrada.Text
  Else
    TotalEntrada = sumaHoras(TotalEntrada, Me.FldEntrada.Text)
 End If
 
   If TotalSalida = "" Then
    TotalSalida = Me.FldSalida.Text
  Else
    TotalSalida = sumaHoras(TotalSalida, Me.FldSalida.Text)
 End If
 
  If TotalInicioComida = "" Then
    TotalInicioComida = Me.FldInicioComida.Text
  Else
    TotalInicioComida = sumaHoras(TotalInicioComida, Me.FldInicioComida.Text)
 End If
 
   If TotalFinComida = "" Then
    TotalFinComida = Me.FldSalidaComida.Text
  Else
    TotalFinComida = sumaHoras(TotalFinComida, Me.FldSalidaComida.Text)
 End If
End Sub

Private Sub GroupFooter1_Format()

  Me.FldTotalEntrada.Text = TotalEntrada
  Me.FldTotalSalida.Text = TotalSalida
  Me.FldTotalInicioComida.Text = TotalInicioComida
  Me.FldTotalFinComida.Text = TotalFinComida
  Me.FldTotalLaboradas.Text = Mid(TotalLaboradas, 1, 5)
  Me.FldTotalExtras.Text = Mid(TotalExtras, 1, 5)
  

End Sub

