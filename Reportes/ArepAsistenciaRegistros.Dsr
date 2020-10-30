VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAsistenciaRegistros 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepAsistenciaRegistros.dsx":0000
End
Attribute VB_Name = "ArepAsistenciaRegistros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalLaboradas As String, TotalExtras As String, TotalEntrada As String, TotalSalida As String, TotalInicioComida As String, TotalFinComida As String, TotalAlmuerzo As String, TotalLaboradas2 As String, TotalExtras2 As String, TotalAlmuerzo2 As String, TotalAusentes As Double, TotalAusentesDpto As Double
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



Private Sub Detail_Format()
  Dim CodEmpleado As String, FechaIni As String, FechaFin As String
  Dim FechaMarca As Date
  

  Me.Field21.ForeColor = &H0&
  Me.Field25.ForeColor = &H0&
  
  If Me.FldFecha.Text <> "" Then
  
      CodEmpleado = Me.Field18.Text
      FechaMarca = Me.FldFechaMarca2.Text
      FechaIni = "#" & Format(FechaMarca, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(FechaMarca, "mm/dd/yyyy") & " 23:59:59#"
      
      Me.Field21.Hyperlink = ""
      Me.Field21.ForeColor = &H0&
          
      If Me.Field21.Text = "0:00" Then

'          Me.Field21.Alignment = ddTXCenter
'          Me.Field21.Text = SimboloNoMarco
          If TieneJustificacion(Me.Field18.Text, Me.FldFechaEntradaInicio, Me.FldFechaEntraFin) = True Then
              Me.Field21.Text = QuienJustifica
              If FrmReportesReloj.ChkLink.Value = 0 Then
               Me.Field21.Hyperlink = CodigoEmpleado
              End If
              Me.Field21.ForeColor = &HC00000
          Else
              Me.Field21.ForeColor = &H0&
              Me.Field21.Hyperlink = ""
          End If
          
      End If
          
          
      Me.Field22.Hyperlink = ""
      Me.Field22.ForeColor = &H0&
      If Me.Field22.Text = "0:00" Then
'          Me.Field21.Alignment = ddTXCenter
'          Me.Field21.Text = SimboloNoMarco
          If TieneJustificacion(Me.Field18.Text, Me.FldFechaSalidaInicio.Text, Me.FldFechaSalidaFin.Text) = True Then
              Me.Field22.Text = QuienJustifica
              If FrmReportesReloj.ChkLink.Value = 0 Then
                Me.Field22.Hyperlink = CodigoEmpleado
              End If
              Me.Field22.ForeColor = &HC00000
              
              
          Else
              Me.Field22.ForeColor = &H0&
              Me.Field22.Hyperlink = ""
          End If
      End If
     
  End If
  
 If TotalLaboradas = "" Then TotalLaboradas = Me.Field25.Text Else TotalLaboradas = sumaHoras(TotalLaboradas, Me.Field25.Text)
 If TotalExtras = "" Then TotalExtras = Me.Field26.Text Else TotalExtras = sumaHoras(TotalExtras, Me.Field26.Text)
' If TotalAlmuerzo = "" Then TotalAlmuerzo = Me.Field27.Text Else TotalAlmuerzo = sumaHoras(TotalAlmuerzo, Me.Field27.Text)
 If TotalLaboradas2 = "" Then TotalLaboradas2 = Me.Field25.Text Else TotalLaboradas2 = sumaHoras(TotalLaboradas2, Me.Field25.Text)
 If TotalExtras2 = "" Then TotalExtras2 = Me.Field26.Text Else TotalExtras2 = sumaHoras(TotalExtras2, Me.Field26.Text)
' If TotalAlmuerzo2 = "" Then TotalAlmuerzo2 = Me.Field27.Text Else TotalAlmuerzo2 = sumaHoras(TotalAlmuerzo2, Me.Field27.Text)
' If TotalSalida = "" Then TotalSalida = Me.FldSalida.Text Else TotalSalida = sumaHoras(TotalSalida, Me.FldSalida.Text)
' If TotalInicioComida = "" Then TotalInicioComida = Me.FldInicioComida.Text Else TotalInicioComida = sumaHoras(TotalInicioComida, Me.FldInicioComida.Text)
' If TotalFinComida = "" Then TotalFinComida = Me.FldSalidaComida.Text Else TotalFinComida = sumaHoras(TotalFinComida, Me.FldSalidaComida.Text)

End Sub
Private Sub ActiveReport_ReportStart()
 QueReporte = "ASISTENCIA X DIA"

 MDIPrimero.DtaEmpresa.Refresh
 Me.LblEmpresa.Caption = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
 Me.LblEmpresa1.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion")
 Me.LblEmpresa2.Caption = "RUC: " & MDIPrimero.DtaEmpresa.Recordset("NumeroRuc")
 Me.LblGenerado.Caption = Format(Now, "mmmm dd,yyyy")
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
Private Sub GroupFooter1_Format()
  Me.FldTotalLaboradas.Text = Mid(TotalLaboradas, 1, 5)
  Me.FldTotalExtras.Text = Mid(TotalExtras, 1, 5)
End Sub

Private Sub GroupHeader1_Format()
'Me.LblFecha.Caption = Format(Me.FldFecha.Text, "Long Date")

TotalHorasExtras = 0
TotalHorasTrabajadas = 0
TotalMinutosTrabajados = 0

 TotalLaboradas = ""
 TotalExtras = ""
 TotalAlmuerzo = ""
 
 TotalAusentes = 0
End Sub
Private Sub ActiveReport_hyperLink(ByVal Button As Integer, Link As String)
    If Button <> 1 Then
        Link = ""
    End If
End Sub
