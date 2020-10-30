VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAsistenciaAusencia 
   Caption         =   "Reporte de Asistencia y Ausencia por dia"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepAsistenciaAusencia.dsx":0000
End
Attribute VB_Name = "ArepAsistenciaAusencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalLaboradas As String, TotalExtras As String, TotalEntrada As String, TotalSalida As String, TotalInicioComida As String, TotalFinComida As String, TotalAlmuerzo As String, TotalLaboradas2 As String, TotalExtras2 As String, TotalAlmuerzo2 As String, TotalAusentes As Double, TotalAusentesDpto As Double
Dim CodigoEmpleado As String


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
Dim CodigoEmpleado As String, CodigoCuentaHasta As String
Dim rpt As Object, FechaIni As String, FechaFin As String
Dim fPreview As New FrmPreview


   If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then
'          ArepAuxiliar.LblRangoFecha = "Desde " & FrmReportes.DTFecha1.Value & " Hasta " & FrmReportes.DTFecha2.Value
'          ArepAuxiliar.LblFecha = Format(Now, "dd/mm/yyyy")
'          ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
'
'
'
'
'            CodigoCuentaDesde = LeeCadena(Link, 1)
'            CodigoCuentaHasta = CodigoCuentaDesde
'
'          sql = "SELECT Transacciones.CodCuentas,  MAX(Transacciones.FechaTransaccion) AS FechaTransaccion, MAX(Cuentas.DescripcionCuentas) AS DescripcionCuentas, MAX(Transacciones.NPeriodo) AS NPeriodo,MAX(Transacciones.NTransaccion) AS NTransaccion, MAX(Transacciones.NumeroMovimiento) AS NumeroMovimiento, MAX(Transacciones.VoucherNo) AS VoucherNo, MAX(Transacciones.DescripcionMovimiento) AS DescripcionMovimiento, MAX(Transacciones.Clave) AS Clave, SUM(Transacciones.Debito) AS Debito, SUM(Transacciones.Credito) AS Credito, MAX(Transacciones.FacturaNo) AS FacturaNo, MAX(Transacciones.ChequeNo) AS ChequeNo, MAX(Transacciones.Fuente) AS Fuente, MAX(Cuentas.TipoCuenta) AS TipoCuenta, SUM(Transacciones.Debito + Transacciones.Credito) As Saldo FROM  Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas WHERE  (Transacciones.FechaTransaccion <= CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) GROUP BY Transacciones.CodCuentas " & _
'                                               "HAVING (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (SUM(Transacciones.Debito + Transacciones.Credito) <> 0) ORDER BY Transacciones.CodCuentas"
'
''         SQL = "SELECT Transacciones.CodCuentas, Cuentas.DescripcionCuentas, Transacciones.NPeriodo, Transacciones.NTransaccion, Transacciones.NumeroMovimiento, Transacciones.VoucherNo, Transacciones.DescripcionMovimiento, Transacciones.Clave, Transacciones.TCambio, Transacciones.FacturaNo, Transacciones.ChequeNo, Transacciones.Fuente, Transacciones.FechaTransaccion, Cuentas.TipoCuenta, Transacciones.TCambio AS Expr1, Transacciones.TCambio * Transacciones.Debito AS Debito, Transacciones.TCambio * Transacciones.Credito AS Credito FROM Cuentas INNER JOIN Transacciones ON Cuentas.CodCuentas = Transacciones.CodCuentas  " & _
''                     "WHERE (Transacciones.CodCuentas BETWEEN '" & CodigoCuentaDesde & "' AND '" & CodigoCuentaHasta & "') AND (Transacciones.FechaTransaccion BETWEEN CONVERT(DATETIME,'" & Format(FrmReportes.DTFecha1.Value, "yyyy-mm-dd") & "', 102) AND CONVERT(DATETIME, '" & Format(FrmReportes.DTFecha2.Value, "yyyy-mm-dd") & "', 102)) ORDER BY Transacciones.CodCuentas, Transacciones.FechaTransaccion, Transacciones.NTransaccion"
'               ArepAuxiliar.DataControl1.ConnectionString = ConexionReporte
'      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
'      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
         
         CodigoEmpleado = Link
         sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid  WHERE (((Userinfo.Userid)='" & CodigoEmpleado & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & "))"
         Set rpt = New ArepJustificacion
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
        
    End If


End Sub


Private Sub ActiveReport_ReportStart()

 QueReporte = "ASISTENCIA Y AUSENCIA"
 MDIPrimero.DtaEmpresa.Refresh
 Me.LblEmpresa.Caption = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
 Me.LblEmpresa1.Caption = MDIPrimero.DtaEmpresa.Recordset("Direccion")
 Me.LblEmpresa2.Caption = "RUC: " & MDIPrimero.DtaEmpresa.Recordset("NumeroRuc")
 RutaLogo = ""
 If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("RutaLogo")) Then
   RutaLogo = MDIPrimero.DtaEmpresa.Recordset("RutaLogo")
 End If
 Me.LblFechaImpreso.Caption = Format(Now, "DD/MM/YYYY")
 Me.LblRangoFecha.Caption = "Impreso desde: " & FrmReportesReloj.DTPFechaIni.Value & " Hasta " & FrmReportesReloj.DTFechaFin.Value
 
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

Private Sub Detail_Format()
  Dim CodEmpleado As String, FechaIni As String, FechaFin As String
  

  Me.Field21.ForeColor = &H0&
  Me.Field25.ForeColor = &H0&
  Me.Field21.Hyperlink = ""
  Me.Field25.Hyperlink = ""
  
   
  
  If Me.FldFecha.Text <> "" Then
      If Me.Field21.Text = "0:00" Then
        If Me.Field22.Text = "0:00" Then
           If Me.Field23.Text = "0:00" Then
              If Me.Field23.Text = "0:00" Then
                CodEmpleado = Me.Field16.Text
                FechaIni = "#" & Format(Me.FldFecha.Text, "mm/dd/yyyy") & "#"
                FechaFin = "#" & Format(Me.FldFecha.Text, "mm/dd/yyyy") & " 23:59:59#"
                'FrmReportesReloj.AdoConsulta.RecordSource = "SELECT UserLeave.*, Format([UserLeave].[BeginTime],'dd/mm/yyyy') AS FechaIni, Format(UserLeave.EndTime,'dd/mm/yyyy') AS FechaFin From UserLeave WHERE (((Format([UserLeave].[BeginTime],'dd/mm/yyyy'))<=" & FechaIni & ") AND ((Format(UserLeave.EndTime,'dd/mm/yyyy'))>=" & FechaIni & ") AND ((UserLeave.Userid)='" & CodEmpleado & "'))"
                FrmReportesReloj.AdoConsulta.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid  " & _
                                                            "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & ")) "

                FrmReportesReloj.AdoConsulta.Refresh
                If Not FrmReportesReloj.AdoConsulta.Recordset.EOF Then
                  
                  If Not IsNull(FrmReportesReloj.AdoConsulta.Recordset("Classname")) Then
                   Me.Field21.Text = FrmReportesReloj.AdoConsulta.Recordset("Classname")
                   Me.Field25.Text = FrmReportesReloj.AdoConsulta.Recordset("Classname")
                  End If
                  If FrmReportesReloj.ChkLink.Value = 0 Then
                  Me.Field21.Hyperlink = CodigoEmpleado
                  Me.Field25.Hyperlink = CodigoEmpleado
                  End If
                  Me.Field25.ForeColor = &HC00000
                  Me.Field21.ForeColor = &HC00000
                Else
                  Me.Field21.Text = "AUSENTE"
                  Me.Field25.Text = "AUSENTE"
                  TotalAusentes = TotalAusentes + 1
                  TotalAusentesDpto = TotalAusentesDpto + 1
                End If
    
              End If
           End If
        Else
                CodEmpleado = Me.Field16.Text
                FechaIni = "#" & Format(Me.FldFecha.Text, "mm/dd/yyyy") & "#"
                FechaFin = "#" & Format(Me.FldFecha.Text, "mm/dd/yyyy") & " 23:59:59#"
                'FrmReportesReloj.AdoConsulta.RecordSource = "SELECT UserLeave.*, Format([UserLeave].[BeginTime],'dd/mm/yyyy') AS FechaIni, Format(UserLeave.EndTime,'dd/mm/yyyy') AS FechaFin From UserLeave WHERE (((Format([UserLeave].[BeginTime],'dd/mm/yyyy'))<=" & FechaIni & ") AND ((Format(UserLeave.EndTime,'dd/mm/yyyy'))>=" & FechaIni & ") AND ((UserLeave.Userid)='" & CodEmpleado & "'))"
                FrmReportesReloj.AdoConsulta.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid  " & _
                                                            "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & ")) "

                FrmReportesReloj.AdoConsulta.Refresh
                If Not FrmReportesReloj.AdoConsulta.Recordset.EOF Then
                  If Not IsNull(FrmReportesReloj.AdoConsulta.Recordset("Classname")) Then
                   Me.Field21.Text = FrmReportesReloj.AdoConsulta.Recordset("Classname")
                   Me.Field25.Text = FrmReportesReloj.AdoConsulta.Recordset("Classname")
                  End If
                  If FrmReportesReloj.ChkLink.Value = 0 Then
                  Me.Field21.Hyperlink = CodigoEmpleado
                  Me.Field25.Hyperlink = CodigoEmpleado
                  End If
                  Me.Field25.ForeColor = &HC00000
                  Me.Field21.ForeColor = &HC00000
                End If
         End If
      End If
  End If
  
 If TotalLaboradas = "" Then TotalLaboradas = Me.Field25.Text Else TotalLaboradas = sumaHoras(TotalLaboradas, Me.Field25.Text)
 If TotalExtras = "" Then TotalExtras = Me.Field26.Text Else TotalExtras = sumaHoras(TotalExtras, Me.Field26.Text)
 If TotalAlmuerzo = "" Then TotalAlmuerzo = Me.Field27.Text Else TotalAlmuerzo = sumaHoras(TotalAlmuerzo, Me.Field27.Text)
 If TotalLaboradas2 = "" Then TotalLaboradas2 = Me.Field25.Text Else TotalLaboradas2 = sumaHoras(TotalLaboradas2, Me.Field25.Text)
 If TotalExtras2 = "" Then TotalExtras2 = Me.Field26.Text Else TotalExtras2 = sumaHoras(TotalExtras2, Me.Field26.Text)
 If TotalAlmuerzo2 = "" Then TotalAlmuerzo2 = Me.Field27.Text Else TotalAlmuerzo2 = sumaHoras(TotalAlmuerzo2, Me.Field27.Text)
' If TotalSalida = "" Then TotalSalida = Me.FldSalida.Text Else TotalSalida = sumaHoras(TotalSalida, Me.FldSalida.Text)
' If TotalInicioComida = "" Then TotalInicioComida = Me.FldInicioComida.Text Else TotalInicioComida = sumaHoras(TotalInicioComida, Me.FldInicioComida.Text)
' If TotalFinComida = "" Then TotalFinComida = Me.FldSalidaComida.Text Else TotalFinComida = sumaHoras(TotalFinComida, Me.FldSalidaComida.Text)

End Sub

Private Sub GroupFooter1_Format()
  Me.Field29.Text = Mid(TotalLaboradas2, 1, 5)
  Me.Field30.Text = Mid(TotalExtras2, 1, 5)
  Me.Field31.Text = Mid(TotalAlmuerzo2, 1, 5)
  Me.LblTotalAusentes.Caption = "Total Ausentes: " & TotalAusentes
End Sub

Private Sub GroupFooter2_Format()
  Me.FldTotalLaboradas.Text = Mid(TotalLaboradas, 1, 5)
  Me.FldTotalExtras.Text = Mid(TotalExtras, 1, 5)
  Me.FldTotalAlmuerzo.Text = Mid(TotalAlmuerzo, 1, 5)
  Me.LblTotalAusentesDpto.Caption = "Total Ausentes " & Me.Field28.Text & ": " & TotalAusentesDpto
End Sub

Private Sub GroupHeader1_Format()
Me.LblFecha.Caption = Format(Me.FldFecha.Text, "Long Date")

TotalHorasExtras = 0
TotalHorasTrabajadas = 0
TotalMinutosTrabajados = 0

 TotalLaboradas = ""
 TotalExtras = ""
 TotalAlmuerzo = ""
 
 TotalAusentes = 0
End Sub

Private Sub GroupHeader2_Format()
 TotalLaboradas = 0
 TotalExtras = 0
 TotalAlmuerzo = 0
 
 TotalAusentesDpto = 0
 
End Sub

