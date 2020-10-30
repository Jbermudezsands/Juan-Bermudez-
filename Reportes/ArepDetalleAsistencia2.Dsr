VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepDetalleAsistencia2 
   Caption         =   "Reporte de Asistencia"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepDetalleAsistencia2.dsx":0000
End
Attribute VB_Name = "ArepDetalleAsistencia2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalHorasExtras As String, TotalMinutosTrabajados As Double, TotalHorasTrabajadas As String, SimboloNoMarco As String
Private Sub ActiveReport_ReportStart()

QueReporte = "ASISTENCIA X DIA"

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
   End If
 End If
 
End Sub

Private Sub Detail_Format()
Dim Entrada As String, Salida As String
If Me.Field18.Text <> "" Then
 Entrada = Me.Field18.Text
Else
 Entrada = "00:00"
End If
If Me.Field19.Text <> "" Then
 Salida = Me.Field19.Text
Else
 Salida = "00:00"
End If
 If Entrada <> "12:00:00 a.m." Then
   If Salida <> "12:00:00 a.m." Then
    TotalHorasTrabajadas = sumaHoras(TotalHorasTrabajadas, Me.Field20.Text)
'   TotalHorasTrabajadas = ConvertirSegundosHoras(DateDiff("s", Entrada, Salida)) + TotalHorasTrabajadas
'   TotalMinutosTrabajados = ConvertirSegundosMinutos(DateDiff("s", Entrada, Salida)) + TotalMinutosTrabajados
  End If
 End If
 
 If Me.Field21.Text <> "0:00" And Me.Field21.Text <> "" Then
'  TotalHorasExtras = CDate(Me.Field21.Text) + TotalHorasExtras
   TotalHorasExtras = sumaHoras(TotalHorasExtras, Me.Field21.Text)
 End If
 
 Me.Field18.Hyperlink = ""
 Me.Field18.ForeColor = &H0&
 
     If Me.FldFechaEntradaInicio.Text <> "" And Me.FldFechaEntraFin.Text <> "" Then
     
        If Me.Field18.Text = "12:00:00 a.m." Then
             If TieneJustificacion(Me.Field16.Text, Me.FldFechaEntradaInicio, Me.FldFechaEntraFin) = True Then
                 If FrmReportesReloj.ChkLink.Value = 0 Then
                  Me.Field18.Hyperlink = Field16.Text
                 End If
                 Me.Field18.Text = QuienJustifica
                 Me.Field18.ForeColor = &HC00000
             Else
                   Me.Field18.Alignment = ddTXCenter
                   Me.Field18.Text = SimboloNoMarco
             End If
        End If
     Else
         Me.Field18.Alignment = ddTXCenter
         Me.Field18.Text = SimboloNoMarco
     End If

    If Me.FldFechaEntradaInicio.Text <> "" And Me.FldFechaEntraFin.Text <> "" Then
             If Me.Field19.Text = "12:00:00 a.m." Then
                     If TieneJustificacion(Me.Field16.Text, Me.FldFechaEntradaInicio, Me.FldFechaEntraFin) = True Then
                        If FrmReportesReloj.ChkLink.Value = 0 Then
                         Me.Field19.Hyperlink = Field16.Text
                        End If
                        Me.Field19.Text = QuienJustifica
                        Me.Field19.ForeColor = &HC00000
                     Else
                        Me.Field19.Alignment = ddTXCenter
                        Me.Field19.Text = SimboloNoMarco
                     End If
             End If
     Else
                        Me.Field19.Alignment = ddTXCenter
                        Me.Field19.Text = SimboloNoMarco
     End If
End Sub

Private Sub GroupFooter1_Format()
Me.LblTotalHorasTrabajadas.Caption = Mid(TotalHorasTrabajadas, 1, 5) 'ConvertirS((TotalHorasTrabajadas * 3600) + (TotalMinutosTrabajados * 60))
Me.LblTotalHorasExtra.Caption = Mid(TotalHorasExtras, 1, 5) 'Format(TotalHorasExtras, "hh:mm")
End Sub

Private Sub GroupHeader1_Format()
TotalHorasTrabajadas = 0
TotalHorasExtras = 0
End Sub

