VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAsistenciaSiete 
   Caption         =   "Reporte de Asistencia"
   ClientHeight    =   10980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepAsistenciaSiete.dsx":0000
End
Attribute VB_Name = "ArepAsistenciaSiete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_ReportStart()
Dim FechaInicio As Date, i As Double, Fecha As Date

 FechaInicio = FrmReportesReloj.DTPFechaIni.Value
 Me.LblFechaIni.Caption = FrmReportesReloj.DTPFechaIni.Value
 Fecha = DateAdd("D", i, FechaInicio)
 LblDia1.Caption = Dia(Fecha)

For i = 1 To 6
 Select Case i
  Case 1:
         Me.LblFecha2.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia2.Caption = Dia(Fecha)
  Case 2:
         Me.LblFecha3.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia3.Caption = Dia(Fecha)
  Case 3:
         Me.LblFecha4.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia4.Caption = Dia(Fecha)
  Case 4:
         Me.LblFecha5.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia5.Caption = Dia(Fecha)
  Case 5:
         Me.LblFecha6.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia6.Caption = Dia(Fecha)
  Case 6:
         Me.LblFecha7.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia7.Caption = Dia(Fecha)
 End Select
Next


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
 
 Me.LblRango.Caption = "Del " & FrmReportesReloj.DTPFechaIni.Value & " Al " & FrmReportesReloj.DTFechaFin.Value
 
  If Not IsNull(MDIPrimero.DtaEmpresa.Recordset("MembreteLogo")) Then
   If MDIPrimero.DtaEmpresa.Recordset("MembreteLogo") = True Then
      Me.Logo.Width = Me.LblEmpresa.Width
      Me.Logo.Height = 700
      Me.PageSettings.TopMargin = 100
      Me.LblEmpresa.Top = 1000
      Me.LblEmpresa1.Top = 1300
      Me.LblEmpresa2.Top = 1550
      Me.Label15.Top = 1800
      Me.LblRango.Top = 2100
   End If
 End If

End Sub

