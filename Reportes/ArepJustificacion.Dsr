VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepJustificacion 
   Caption         =   "REPORTE DE JUSTIFICACIONES"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepJustificacion.dsx":0000
End
Attribute VB_Name = "ArepJustificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TotalHorasJustificadas As String, TotalMinutosTraba As Double, TotalHorasTrabajadas As Double, SimboloNoMarco As String, MembreteLogo As Boolean

Private Sub ActiveReport_ReportEnd()
   Unload SubDetalle.object
   Set SubDetalle.object = Nothing
End Sub

Private Sub ActiveReport_ReportStart()

 Set SubDetalle.object = New ArepSubJustificacion
 
      If FrmReportesReloj.ChkTodosDptos.Value = 1 Then
         Me.GroupHeader1.Visible = True
         Me.GroupFooter1.Visible = True
      ElseIf FrmReportesReloj.DBDptoIni.Text <> "" And FrmReportesReloj.DBDptoFin.Text <> "" Then
         Me.GroupHeader1.Visible = True
         Me.GroupFooter1.Visible = True
      Else
         Me.GroupHeader1.Visible = False
         Me.GroupFooter1.Visible = Fals
      End If
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
Dim FechaIni As String, FechaFin As String, CodEmpleado As String


    FechaIni = "#" & Format(FrmReportesReloj.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
    FechaFin = "#" & Format(FrmReportesReloj.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
    CodEmpleado = Me.FldCodEmpleado.Text
      
    SubDetalle.object.DataControl1.ConnectionString = ConexionReloj
    SubDetalle.object.DataControl1.Source = "SELECT Reportes.Campo1 AS Userid, Reportes.Campo2 AS Name, Reportes.Campo3 AS DeptName, Reportes.CampoFecha1 AS BeginTime, Reportes.CampoFecha2 AS EndTime, Reportes.Campo4 AS Classname From Reportes WHERE (((Reportes.Campo1)='" & CodEmpleado & "'))"

'    SubDetalle.object.DataControl1.Source = "SELECT DISTINCT Userinfo.Userid, Userinfo.Name, UserLeave.BeginTime, UserLeave.EndTime, LeaveClass.Classname FROM LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid  " & _
                                          "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserLeave.BeginTime)>=" & FechaIni & ") AND ((UserLeave.EndTime)<=" & FechaFin & "))"
  

End Sub

Private Sub GroupFooter1_Format()
Me.LblDescripcion.Caption = "Total " & Me.Field17.Text
Me.LblTotalHorasDpto.Caption = Mid(SubDetalle.object.TotalHorasDepartamento, 1, 5)
SubDetalle.object.TotalHorasDepartamento = ""
End Sub

Private Sub GroupFooter2_Format()
TotalHorasJustificadas = SubDetalle.object.TotalHoraEmpleado
Me.LblTotalHoras.Caption = Mid(TotalHorasJustificadas, 1, 5)
SubDetalle.object.TotalHoraEmpleado = ""
End Sub

