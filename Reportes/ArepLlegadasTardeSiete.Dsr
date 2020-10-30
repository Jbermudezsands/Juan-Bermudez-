VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ArepLLegadasTardeSiete 
   Caption         =   "Reporte de Llegadas Tarde Siete"
   ClientHeight    =   11010
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   20280
   StartUpPosition =   3  'Windows Default
   _ExtentX        =   35878
   _ExtentY        =   19420
   SectionData     =   "ArepLlegadasTardeSiete.dsx":0000
End
Attribute VB_Name = "ArepLLegadasTardeSiete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public HorasTardeLunes As String, HorasTardeMartes As String, HorasTardeMiercoles As String, HorasTardeJueves As String, HorasTardeViernes As String, HorasTardeSabado As String, HorasTardeDomingo As String

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
      Me.Label23.Top = 2350
   End If
 End If
 
End Sub

Private Sub Detail_Format()
 Dim TotalHorasTarde As String
 '==================================================================================================
 If HorasTardeLunes = "" Then
    HorasTardeLunes = Me.FldTardeLunes.Text
  Else
    HorasTardeLunes = sumaHoras(HorasTardeLunes, Me.FldTardeLunes.Text)
 End If
 
 If HorasTardeMartes = "" Then
    HorasTardeMartes = Me.FldTardeMartes.Text
  Else
    HorasTardeMartes = sumaHoras(HorasTardeMartes, Me.FldTardeMartes.Text)
 End If
 
  If HorasTardeMiercoles = "" Then
    HorasTardeMiercoles = Me.FldTardeMiercoles.Text
  Else
    HorasTardeMiercoles = sumaHoras(HorasTardeMiercoles, Me.FldTardeMiercoles.Text)
 End If

  If HorasTardeJueves = "" Then
    HorasTardeJueves = Me.FldTardeJueves.Text
  Else
    HorasTardeJueves = sumaHoras(HorasTardeJueves, Me.FldTardeJueves.Text)
 End If
 
  If HorasTardeViernes = "" Then
    HorasTardeViernes = Me.FldTardeViernes.Text
  Else
    HorasTardeViernes = sumaHoras(HorasTardeViernes, Me.FldTardeViernes.Text)
 End If

 If HorasTardeSabado = "" Then
    HorasTardeSabado = Me.FldTardeSabado.Text
  Else
    HorasTardeSabado = sumaHoras(HorasTardeSabado, Me.FldTardeSabado.Text)
 End If

 If HorasTardeDomingo = "" Then
    HorasTardeDomingo = Me.FldTardeDomingo.Text
  Else
    HorasTardeDomingo = sumaHoras(HorasTardeDomingo, Me.FldTardeDomingo.Text)
 End If
 
 
  TotalHorasTarde = sumaHoras(Me.FldTardeLunes.Text, Me.FldTardeMartes.Text)
  TotalHorasTarde = sumaHoras(Me.FldTardeMiercoles.Text, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(Me.FldTardeJueves.Text, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(Me.FldTardeViernes.Text, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(Me.FldTardeSabado.Text, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(Me.FldTardeDomingo.Text, TotalHorasTarde)
  Me.FldTardeTotal.Text = Mid(TotalHorasTarde, 1, 5)
End Sub

Private Sub GroupFooter1_Format()
  Dim TotalHorasTarde As String
  
  TotalHorasTarde = "00:00"
  
  Me.FldTotalTardeLunes.Text = Mid(HorasTardeLunes, 1, 5)
  Me.FldTotalTardeMartes.Text = Mid(HorasTardeMartes, 1, 5)
  Me.FldTotalTardeMiercoles.Text = Mid(HorasTardeMiercoles, 1, 5)
  Me.FldTotalTardeJueves.Text = Mid(HorasTardeJueves, 1, 5)
  Me.FldTotalTardeViernes.Text = Mid(HorasTardeViernes, 1, 5)
  Me.FldTotalTardeSabado.Text = Mid(HorasTardeSabado, 1, 5)
  Me.FldTotalTardeDomingo.Text = Mid(HorasTardeDomingo, 1, 5)
  
  TotalHorasTarde = sumaHoras(HorasTardeLunes, HorasTardeMartes)
  TotalHorasTarde = sumaHoras(HorasTardeMiercoles, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(HorasTardeJueves, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(HorasTardeViernes, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(HorasTardeSabado, TotalHorasTarde)
  TotalHorasTarde = sumaHoras(HorasTardeDomingo, TotalHorasTarde)
  Me.FldTotalGeneralTarde.Text = Mid(TotalHorasTarde, 1, Len(TotalHorasTarde) - 3)
End Sub

Private Sub GroupHeader1_Format()
 
  HorasTardeLunes = ""
  HorasTardeMartes = ""
  HorasTardeMiercoles = ""
  HorasTardeJueves = ""
  HorasTardeViernes = ""
  HorasTardeSabado = ""
  HorasTardeDomingo = ""
End Sub
