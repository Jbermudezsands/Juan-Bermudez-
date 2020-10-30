VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAusencia 
   Caption         =   "Ausencia de Empleados"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepAusencia.dsx":0000
End
Attribute VB_Name = "ArepAusencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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
 Me.LblRangoFecha.Caption = "Impreso Desde:" & Format(FrmReportesReloj.DTPFechaIni.Value, "dd/mm/yyyy") & " Hasta " & Format(FrmReportesReloj.DTFechaFin.Value, "dd/mm/yyyy")
 
 If (Dir(RutaLogo) <> "") Then
    Me.Logo.Picture = LoadPicture(RutaLogo)
 End If
 
    If FrmReportesReloj.ChkTodosDptos.Value = 0 Then
     If FrmReportesReloj.DBDptoIni.Text <> "" Or FrmReportesReloj.DBDptoFin.Text <> "" Then
             Me.GroupHeader1.Visible = True
             Me.GroupFooter1.Visible = True
     End If
    Else

             Me.GroupHeader1.Visible = True
             Me.GroupFooter1.Visible = True

            
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
      Me.LblTipoJustifica.Hyperlink = ""
      Me.LblTipoJustifica.ForeColor = &H0&
          


'          Me.Field21.Alignment = ddTXCenter
'          Me.Field21.Text = SimboloNoMarco
          If TieneJustificacion(Me.Field16.Text, Me.FldFechaEntradaInicio, Me.FldFechaEntraFin) = True Then
'              Me.LblTipoJustifica.Text = "JUSTIFICA"
              If FrmReportesReloj.ChkLink.Value = 0 Then
               Me.LblTipoJustifica.Hyperlink = Me.Field16.Text
              End If
              Me.LblTipoJustifica.ForeColor = &HC00000
              Me.LblTipoJustifica.Caption = TipoJustificacion(Me.Field16.Text, Me.FldFechaEntradaInicio, Me.FldFechaEntraFin)
          Else
              Me.LblTipoJustifica.ForeColor = &H0&
              Me.LblTipoJustifica.Hyperlink = ""
              Me.LblTipoJustifica.Caption = ""
          End If
          

End Sub

