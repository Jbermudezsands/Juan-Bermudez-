VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAsistenciaDpto 
   Caption         =   "Asistencia x Departamentos"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   26882
   _ExtentY        =   19368
   SectionData     =   "ArepAsistenciaDpto.dsx":0000
End
Attribute VB_Name = "ArepAsistenciaDpto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalHorasExtras As Date, TotalMinutosTrabajados As Double, TotalHorasTrabajadas As Double, SimboloNoMarco As String
Public TotalLaboradas As String, TotalExtras As String

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
 
 Me.LblRango.Caption = "Impreso desde: " & FrmReportesReloj.DTPFechaIni.Value & " Hasta: " & FrmReportesReloj.DTFechaFin.Value
 
 If Quien = "Ausencia" Then
    Me.Label15.Caption = "REPORTE DE AUSENCIA"
 ElseIf Quien = "Marcas Incompletas" Then
    Me.Label15.Caption = "REPORTE DE MARCAS INCOMPLETAS"
 
 End If
 
 SimboloNoMarco = "N/M"
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("SimboloNoMarco")) Then
    If MDIPrimero.DtaEmpresa.Recordset("SimboloNoMarco") <> "" Then
        SimboloNoMarco = MDIPrimero.DtaEmpresa.Recordset("SimboloNoMarco")
    End If
 End If
 
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("MembreteLogo")) Then
   If MDIPrimero.DtaEmpresa.Recordset("MembreteLogo") = True Then
      Me.Logo.Width = 9500
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
Dim Entrada As Date, Salida As Date
Dim Hora As Date

Hora = "12:00:00 a.m."

If Me.Field18.Text <> "" Then
 Entrada = Me.Field18.Text
End If
If Me.Field19.Text <> "" Then
 Salida = Me.Field19.Text
End If

 If Entrada <> "12:00:00 a.m." And Salida <> "12:00:00 a.m." Then
   TotalHorasTrabajadas = ConvertirSegundosHoras(DateDiff("s", Entrada, Salida)) + TotalHorasTrabajadas
   TotalMinutosTrabajados = ConvertirSegundosMinutos(DateDiff("s", Entrada, Salida)) + TotalMinutosTrabajados
   
 End If
 
 If Me.Field21.Text <> "0:00" And Me.Field21.Text <> "" Then
'  TotalHorasExtras = Me.Field21.Text + TotalHorasExtras
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
 
 If TotalLaboradas = "" Then TotalLaboradas = Me.Field20.Text Else TotalLaboradas = sumaHoras(TotalLaboradas, Me.Field20.Text)
 If TotalExtras = "" Then TotalExtras = Me.Field21.Text Else TotalExtras = sumaHoras(TotalExtras, Me.Field21.Text)
 
End Sub

Private Sub GroupFooter1_Format()
Me.LblTotalHorasTrabajadas.Caption = ConvertirS((TotalHorasTrabajadas * 3600) + (TotalMinutosTrabajados * 60))
Me.LblTotalHorasExtra.Caption = Format(TotalHorasExtras, "hh:mm")
End Sub

Private Sub GroupFooter2_Format()
  Me.FldTotalLaboradas.Caption = Mid(TotalLaboradas, 1, 5)
  Me.FldTotalExtras.Caption = Mid(TotalExtras, 1, 5)
End Sub

Private Sub GroupHeader1_Format()
Me.LblFecha.Caption = Format(Me.FldFecha.Text, "Long Date")

TotalHorasExtras = 0
TotalHorasTrabajadas = 0
TotalMinutosTrabajados = 0
End Sub

Private Sub GroupHeader2_Format()
 TotalLaboradas = ""
 TotalExtras = ""
End Sub

