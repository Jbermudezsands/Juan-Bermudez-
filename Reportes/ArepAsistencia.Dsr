VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAsistencia 
   Caption         =   "Reporte de Asistencias"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepAsistencia.dsx":0000
End
Attribute VB_Name = "ArepAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalHorasExtras As Date, TotalMinutosTrabajados As Double, TotalHorasTrabajadas As Double, SimboloNoMarco As String, MembreteLogo As Boolean
Public CodigoEmpleado As String

Private Sub ActiveReport_FetchData(EOF As Boolean)
    If Not EOF Then
    'Gets the current records SupplierID
        If Not IsNull(Me.DataControl1.Recordset.Fields("CodEmpleado")) Then
          CodigoEmpleado = Me.DataControl1.Recordset.Fields("CodEmpleado")
        Else
          CodigoEmpleado = ""
        End If
    End If
End Sub

Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    If Button <> 1 Then
        Link = ""
    End If
End Sub

Private Sub ActiveReport_ReportStart()
  QueReporte = "ASISTENCIA X DIA"
 '******************************************************************************
 '//////BUSCO LA CONFIGURACION GENERAL /////////////////////////////////////////
 '*****************************************************************************
 MDIPrimero.DtaEmpresa.Refresh
 Me.LblEmpresa.Caption = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
 Me.LblEmpresa1.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion")
 Me.LblEmpresa2.Caption = "RUC: " & MDIPrimero.DtaEmpresa.Recordset("NumeroRuc")
 Me.LblRango.Caption = "Impreso desde: " & FrmReportesReloj.DTPFechaIni.Value & " Hasta: " & FrmReportesReloj.DTFechaFin.Value
 Me.LblNombreReporte.Caption = FrmReportesReloj.CmbReportes.Text
 RutaLogo = ""
 
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("RutaLogo")) Then
   RutaLogo = MDIPrimero.DtaEmpresa.Recordset("RutaLogo")
 End If
 Me.LblFechaImpreso.Caption = Format(Now, "DD/MM/YYYY")
 
 If DirRutaLogo <> "" Then
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
      Me.LblNombreReporte.Top = 1800
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
   TotalHorasTrabajadas = ConvertirSegundosHoras(DateDiff("s", Entrada, Salida)) + TotalHorasTrabajadas
   TotalMinutosTrabajados = ConvertirSegundosMinutos(DateDiff("s", Entrada, Salida)) + TotalMinutosTrabajados
  End If
 End If
 
 If Me.Field21.Text <> "0:00" And Me.Field21.Text <> "" Then
'  TotalHorasExtras = CDate(Me.Field21.Text) + TotalHorasExtras
 End If
 
 Me.Field18.Hyperlink = ""
 Me.Field18.ForeColor = &H0&
 
 
 If Me.Field18.Text = "12:00:00 a.m." Then
   Me.Field18.Alignment = ddTXCenter
   Me.Field18.Text = SimboloNoMarco
   If TieneJustificacion(Me.Field16.Text, Me.FldFechaEntradaInicio, Me.FldFechaEntraFin) = True Then
     If FrmReportesReloj.ChkLink.Value = 0 Then
        Me.Field18.Hyperlink = CodigoEmpleado
     End If
        Me.Field18.Text = QuienJustifica
        Me.Field18.ForeColor = &HC00000
   Else
'    Me.Field18.ForeColor = &H0&
'    Me.Field18.Hyperlink = CodigoEmpleado
    Me.Field18.ForeColor = &H0&
    Me.Field18.Text = SimboloNoMarco
   End If

 End If

Me.Field19.Hyperlink = ""
Me.Field19.ForeColor = &H0&

 If Me.Field19.Text = "12:00:00 a.m." Then
   Me.Field19.Alignment = ddTXCenter
   
   If TieneJustificacion(Me.Field16.Text, Me.FldFechaSalidaInicio.Text, Me.FldFechaSalidaFin.Text) = True Then
    If FrmReportesReloj.ChkLink.Value = 0 Then
        Me.Field19.Hyperlink = CodigoEmpleado
    End If
        Me.Field19.Text = QuienJustifica
        Me.Field19.ForeColor = &HC00000
   Else
    Me.Field19.ForeColor = &H0&
    Me.Field19.Text = SimboloNoMarco
   End If
 End If

End Sub

Private Sub GroupFooter1_Format()
Me.LblTotalHorasTrabajadas.Caption = ConvertirS((TotalHorasTrabajadas * 3600) + (TotalMinutosTrabajados * 60))
Me.LblTotalHorasExtra.Caption = Format(TotalHorasExtras, "hh:mm")
End Sub

Private Sub GroupHeader1_Format()
 If FrmReportesReloj.ChkAcumulado.Value = 0 Then
    Me.LblFecha.Caption = Format(Me.FldFecha.Text, "Long Date")
 Else
   Me.LblFecha.Caption = "Desde  " & FrmReportesReloj.DTPFechaIni.Value & "  Hasta  " & FrmReportesReloj.DTFechaFin.Value
 End If

TotalHorasExtras = 0
TotalHorasTrabajadas = 0
TotalMinutosTrabajados = 0
End Sub

