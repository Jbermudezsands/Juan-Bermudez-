VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepAsistenciaQuince 
   Caption         =   "ActiveReport1"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35772
   _ExtentY        =   19368
   SectionData     =   "ArepAsistenciaQuince.dsx":0000
End
Attribute VB_Name = "ArepAsistenciaQuince"
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

For i = 1 To 14
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
  Case 7:
         Me.LblFecha8.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia8.Caption = Dia(Fecha)
  Case 8:
         Me.LblFecha9.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia9.Caption = Dia(Fecha)
  Case 9:
         Me.LblFecha10.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia10.Caption = Dia(Fecha)
  Case 10:
         Me.LblFecha11.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia11.Caption = Dia(Fecha)
  Case 11:
         Me.LblFecha12.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia12.Caption = Dia(Fecha)
  Case 12:
         Me.LblFecha13.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia13.Caption = Dia(Fecha)
  Case 13:
         Me.LblFecha14.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia14.Caption = Dia(Fecha)
   Case 14:
         Me.LblFecha15.Caption = DateAdd("D", i, FechaInicio)
         Fecha = DateAdd("D", i, FechaInicio)
         LblDia15.Caption = Dia(Fecha)

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

Private Sub Detail_Format()
' If Me.Field19.Text = "00:00" Then
'    Me.Field19.Text = ""
' End If
' If Me.Field20.Text = "00:00" Then
'    Me.Field20.Text = ""
' End If
'  If Me.Field21.Text = "00:00" Then
'    Me.Field21.Text = ""
' End If
'  If Me.Field22.Text = "00:00" Then
'    Me.Field22.Text = ""
' End If
'   If Me.Field23.Text = "00:00" Then
'    Me.Field23.Text = ""
' End If
'   If Me.Field24.Text = "00:00" Then
'    Me.Field24.Text = ""
' End If
'   If Me.Field25.Text = "00:00" Then
'    Me.Field25.Text = ""
' End If
'   If Me.Field26.Text = "00:00" Then
'    Me.Field26.Text = ""
' End If
' If Me.Field27.Text = "00:00" Then
'    Me.Field27.Text = ""
' End If
' If Me.Field28.Text = "00:00" Then
'    Me.Field28.Text = ""
' End If
' If Me.Field29.Text = "00:00" Then
'    Me.Field29.Text = ""
' End If
' If Me.Field30.Text = "00:00" Then
'    Me.Field30.Text = ""
' End If
' If Me.Field32.Text = "00:00" Then
'    Me.Field32.Text = ""
' End If
'  If Me.Field33.Text = "00:00" Then
'    Me.Field33.Text = ""
' End If
'  If Me.Field34.Text = "00:00" Then
'    Me.Field34.Text = ""
' End If
'  If Me.Field35.Text = "00:00" Then
'    Me.Field35.Text = ""
' End If
'  If Me.Field36.Text = "00:00" Then
'    Me.Field36.Text = ""
' End If
'  If Me.Field37.Text = "00:00" Then
'    Me.Field37.Text = ""
' End If
'  If Me.Field38.Text = "00:00" Then
'    Me.Field38.Text = ""
' End If
'  If Me.Field39.Text = "00:00" Then
'    Me.Field39.Text = ""
' End If
'   If Me.Field40.Text = "00:00" Then
'    Me.Field40.Text = ""
' End If
'   If Me.Field41.Text = "00:00" Then
'    Me.Field41.Text = ""
' End If
'   If Me.Field42.Text = "00:00" Then
'    Me.Field42.Text = ""
' End If
'   If Me.Field43.Text = "00:00" Then
'    Me.Field43.Text = ""
' End If
'   If Me.Field44.Text = "00:00" Then
'    Me.Field44.Text = ""
' End If
'   If Me.Field45.Text = "00:00" Then
'    Me.Field45.Text = ""
' End If
'   If Me.Field46.Text = "00:00" Then
'    Me.Field46.Text = ""
' End If
'   If Me.Field47.Text = "00:00" Then
'    Me.Field47.Text = ""
' End If
'   If Me.Field48.Text = "00:00" Then
'    Me.Field48.Text = ""
' End If
End Sub

Private Sub PageHeader_Format()

End Sub
