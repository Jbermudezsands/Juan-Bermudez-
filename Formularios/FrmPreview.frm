VERSION 5.00
Object = "{A45D986F-3AAF-4A3B-A003-A6C53E8715A2}#1.0#0"; "ARVIEW2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmPreview 
   Caption         =   "Form1"
   ClientHeight    =   4500
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   4500
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   360
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin DDActiveReportsViewer2Ctl.ARViewer2 arv 
      Height          =   4335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   7646
      SectionData     =   "FrmPreview.frx":0000
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&Archivo"
      Begin VB.Menu ExportaPDF 
         Caption         =   "&Exporta  PDF"
      End
      Begin VB.Menu ExportaExcel 
         Caption         =   "&Exportar Excel"
      End
   End
End
Attribute VB_Name = "FrmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub arv_hyperLink(ByVal Button As Integer, Link As String)
Dim CodigoEmpleado As String, CodigoCuentaHasta As String
Dim rpt As Object, FechaIni As String, FechaFin As String
Dim fPreview As New FrmPreview



   If InStr(1, Link, "htm", vbTextCompare) = 0 And InStr(1, Link, "mailto", vbTextCompare) = 0 Then

      Select Case QueReporte
      
       Case "ASISTENCIA Y AUSENCIA"
                    FechaIni = "#" & Format(FrmReportesReloj.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
                    FechaFin = "#" & Format(FrmReportesReloj.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
                       
                       CodigoEmpleado = Link
                       sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid  WHERE (((Userinfo.Userid)='" & CodigoEmpleado & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & "))"
                       Set rpt = New ArepJustificacion2
                       rpt.DataControl1.ConnectionString = ConexionEasy
                       rpt.DataControl1.Source = sql
                       fPreview.RunReport rpt
                       fPreview.Show 1
      
      Case "ASISTENCIA X DIA"
                    FechaIni = "#" & Format(FrmReportesReloj.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
                    FechaFin = "#" & Format(FrmReportesReloj.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
                       
                       CodigoEmpleado = Link
                       sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid  WHERE (((Userinfo.Userid)='" & CodigoEmpleado & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & "))"
                       Set rpt = New ArepJustificacion2
                       rpt.DataControl1.ConnectionString = ConexionEasy
                       rpt.DataControl1.Source = sql
                       fPreview.RunReport rpt
                       fPreview.Show 1
      
      
      End Select
         
        
    End If
End Sub



Private Sub ExportaExcel_Click()
Dim xls As New ActiveReportsExcelExport.ARExportExcel
Dim sFile As String
Dim bSave As Boolean

On Error GoTo TipoErrs
   
    Me.CommonDialog.Filter = "Formato Excel (*.xls)| *.xls"
    Me.CommonDialog.ShowSave
'    bSave = Dir(Me.CommonDialog.FileName + ".xls")
    
'    If bSave Then xls.FileName = sFile Else Exit Sub

    sFile = Me.CommonDialog.FileName
    xls.FileName = sFile
    
    If arv.Pages.Count > 0 Then
        xls.Export arv.Pages
    ElseIf Not arv.ReportSource Is Nothing Then
        If arv.ReportSource.Pages.Count > 0 Then
            xls.Export arv.ReportSource.Pages
        End If
    End If
    Set xls = Nothing
    MsgBox "Se ha Exportado el Archivo", vbExclamation, "Zeus Contabilidad"
    
    Exit Sub
TipoErrs:
 MsgBox Err.Description
    
End Sub

Private Sub ExportaPDF_Click()
Dim pdf As New ActiveReportsPDFExport.ARExportPDF
Dim sFile As String
Dim bSave As Boolean

On Error GoTo TipoErrs

    Me.CommonDialog.Filter = "Portable Document Format (*.PDF)| *.PDF"
    Me.CommonDialog.ShowSave
    sFile = Me.CommonDialog.FileName
    
    
    pdf.FileName = sFile
    
    If arv.Pages.Count > 0 Then
        pdf.Export arv.Pages
    ElseIf Not arv.ReportSource Is Nothing Then
        If arv.ReportSource.Pages.Count > 0 Then
            pdf.Export arv.ReportSource.Pages
        End If
    End If
    
    Set pdf = Nothing
    MsgBox "Se ha Exportado el Archivo", vbExclamation, "Zeus Contabilidad"
    
Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Public Sub RunReport(rpt As Object)
    Set arv.ReportSource = rpt
    
    arv.Zoom = 100
    Caption = rpt.Caption
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    arv.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
