VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepJustificacion2 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepJustificacion2.dsx":0000
End
Attribute VB_Name = "ArepJustificacion2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalHoraEmpleado As String, TotalHorasDepartamento As String

Private Sub ActiveReport_ReportStart()
'******************************************************************************
 '//////BUSCO LA CONFIGURACION GENERAL /////////////////////////////////////////
 '*****************************************************************************
 MDIPrimero.DtaEmpresa.Refresh
 Me.LblEmpresa.Caption = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
 Me.LblEmpresa1.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion")
 Me.LblEmpresa2.Caption = "RUC: " & MDIPrimero.DtaEmpresa.Recordset("NumeroRuc")
 RutaLogo = ""
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("RutaLogo")) Then
   RutaLogo = MDIPrimero.DtaEmpresa.Recordset("RutaLogo")
 End If
 Me.LblFechaImpreso.Caption = Format(Now, "DD/MM/YYYY")
 Me.LblDesde.Caption = "Desde: " & FrmReportesReloj.DTPFechaIni.Value & " Hasta: " & FrmReportesReloj.DTFechaFin.Value
 
 If (Dir(RutaLogo) <> "") Then
    Me.Logo.Picture = LoadPicture(RutaLogo)
 End If
  
  SimboloNoMarco = "N/M"
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("SimboloNoMarco")) Then
    If MDIPrimero.DtaEmpresa.Recordset("SimboloNoMarco") <> "" Then
        SimboloNoMarco = MDIPrimero.DtaEmpresa.Recordset("SimboloNoMarco")
    End If
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
      Me.LblDesde.Top = 2000
   End If
 End If

End Sub

Private Sub Detail_Format()
Dim TotalHoras As String

TotalHoras = DateDiff("s", Me.FldBegin.Text, Me.FldEnd.Text)
Me.LblTotalHoras.Caption = Int(TotalHoras / 3600) & ":" & Int((TotalHoras Mod 3600) / 60)

TotalHoraEmpleado = sumaHoras(Me.LblTotalHoras.Caption, TotalHoraEmpleado)
TotalHorasDepartamento = sumaHoras(Me.LblTotalHoras.Caption, TotalHorasDepartamento)

'Me.LblFechaHora.Caption = Format(Me.FldBegin.Text, "dd/mm/yyyy") & "-" & Format(Me.FldEnd.Text, "dd/mm/yyyy")
'Me.LblFechaHora.Caption = Format(Me.FldBegin.Text, "dd/mm/yyyy HH:MM") & " - " & Format(Me.FldEnd.Text, "dd/mm/yyyy HH:MM")
End Sub

Private Sub GroupFooter1_Format()
Me.LblTotalHorasDpto.Caption = TotalHorasDepartamento
End Sub

Private Sub GroupFooter2_Format()
Me.LblTotalHoras.Caption = TotalHorasJustificadas
End Sub

