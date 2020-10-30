VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form FrmDepartamentoReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos"
   ClientHeight    =   6120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   3735
      Begin VB.CommandButton SmartButton1 
         Caption         =   "Pegar"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton SmartButton7 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   5520
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4440
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentoReportes.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentoReportes.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentoReportes.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentoReportes.frx":0CF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3720
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentoReportes.frx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentoReportes.frx":24D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   9128
      _Version        =   393217
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      SingleSel       =   -1  'True
      ImageList       =   "ImageList3"
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSAdodcLib.Adodc DtaConsulta 
      Height          =   375
      Left            =   240
      Top             =   6480
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "DtaConsulta"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc DtaGrupos 
      Height          =   375
      Left            =   240
      Top             =   7200
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Grupos"
      Caption         =   "DtaGrupos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmDepartamentoReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MDIPrimero.Skin1.ApplySkin hWnd


Dim NodX As Node
Dim Relatives As String, RelationsShips As String
Dim LLave As String, Texto As String, Imagen1 As Integer
Dim Imagen2 As Integer

With Me.DtaGrupos
   '.DatabaseName = Ruta
   .ConnectionString = ConexionEasy
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = ConexionEasy
End With


i = 1
 ReDim MatrizCuentas(100)
' Me.DtaGrupos.RecordSource = "SELECT Grupos.KeyGrupo, Grupos.KeyGrupoSuperior, Grupos.Child, Grupos.DescripcionGrupo, Grupos.Imagen1, Grupos.Imagen2 From Grupos ORDER BY Grupos.KeyGrupo"
 Me.DtaGrupos.RecordSource = "SELECT Dept.Deptid AS KeyGrupo, Dept.DeptName AS DescripcionGrupo, Dept.SupDeptid AS KeyGrupoSuperior FROM Dept ORDER BY Dept.Deptid"
 Me.DtaGrupos.Refresh
 Do While Not Me.DtaGrupos.Recordset.EOF
   If Me.DtaGrupos.Recordset("KeyGrupoSuperior") <> 0 Then
    Relatives = "A" & Me.DtaGrupos.Recordset("KeyGrupoSuperior")
   Else
     Relatives = "A"
   End If

'   If Not IsNull(Me.DtaGrupos.Recordset("KeyGrupoSuperior")) Then
'     RelationsShips = "4" & Me.DtaGrupos.Recordset("KeyGrupoSuperior")
'   Else
'     RelationsShips = ""
'   End If
   RelationsShips = "4"

   LLave = "A" & Me.DtaGrupos.Recordset("KeyGrupo")
   Texto = Me.DtaGrupos.Recordset("DescripcionGrupo")
   Imagen1 = 4
   Imagen2 = 3
   
   If Relatives = "A" Then
     Set NodX = Me.TreeView1.Nodes.Add(, , LLave, Texto, Imagen1, Imagen2)
   Else
     Set NodX = Me.TreeView1.Nodes.Add(Relatives, RelationsShips, LLave, Texto, Imagen1, Imagen2)
   End If
   
  Me.DtaGrupos.Recordset.MoveNext
 Loop



KeyPrincipal = "A"
Me.TreeView1.Nodes(Me.TreeView1.Nodes.Count).EnsureVisible
NodoBase = True

End Sub

Private Sub SmartButton1_Click()
Select Case Quien
  Case "DptoIni"
      FrmReportesReloj.DBDptoIni.Text = Me.TreeView1.SelectedItem
  Case "DptoFin"
      FrmReportesReloj.DBDptoFin.Text = Me.TreeView1.SelectedItem
  
End Select

Unload Me
End Sub

Private Sub SmartButton7_Click()
Unload Me
End Sub

