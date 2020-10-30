VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepSalidaAnticipada 
   Caption         =   "Salidas Anticipadas"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepSalidaAnticipada.dsx":0000
End
Attribute VB_Name = "ArepSalidaAnticipada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
  QueReporte = "ASISTENCIA X DIA"
 
 MDIPrimero.DtaEmpresa.Refresh
 Me.LblEmpresa.Caption = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
 Me.LblEmpresa1.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion")
 Me.LblEmpresa2.Caption = "RUC: " & MDIPrimero.DtaEmpresa.Recordset("NumeroRuc")
 RutaLogo = ""
 
 Me.LblRango.Caption = "Impreso desde: " & FrmReportesReloj.DTPFechaIni.Value & " Hasta: " & FrmReportesReloj.DTFechaFin.Value
 
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
 
         If FrmReportesReloj.DBDptoIni.Text = "" And FrmReportesReloj.DBDptoFin.Text = "" Then
          Me.GroupFooter2.Visible = False
          Me.GroupHeader2.Visible = False
         End If
         
End Sub

Private Sub Detail_Format()
   Me.FldJustificacion.Hyperlink = ""
   Me.FldJustificacion.Text = ""
   
   If TieneJustificacion(Me.Field16.Text, Me.FldFechaEntradaInicio.Text, Me.FldFechaEntradaFin.Text) = True Then
    If FrmReportesReloj.ChkLink.Value = 0 Then
      Me.FldJustificacion.Hyperlink = Me.Field16.Text
    End If
    Me.FldJustificacion.Text = QuienJustifica
    Me.FldJustificacion.ForeColor = &HC00000
   Else
    Me.FldJustificacion.ForeColor = &H0&
    Me.FldJustificacion.Text = ""
   End If
End Sub

Private Sub GroupHeader1_Format()
Me.LblFecha.Caption = Format(Me.FldFecha.Text, "Long Date")
End Sub

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

