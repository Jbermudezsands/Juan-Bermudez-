VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{080026CA-5CAE-11D6-82C2-000021B74250}#16.0#0"; "vbskfree.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmControlesReloj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Controles Personalizados"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   7800
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7815
      TabIndex        =   45
      Top             =   -120
      Width           =   7815
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "CONTROLES PERSONALIZADOS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2040
         TabIndex        =   46
         Top             =   360
         Width           =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   240
         Picture         =   "FrmControles.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   240
      Top             =   6720
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "AdoEmpleados"
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
   Begin MSAdodcLib.Adodc AdoDetalleNominas 
      Height          =   375
      Left            =   0
      Top             =   8520
      Visible         =   0   'False
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
      Caption         =   "AdoDetalleNominas"
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
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "Aceptar"
      DownPicture     =   "FrmControles.frx":09DD
      Height          =   735
      Left            =   5640
      MouseIcon       =   "FrmControles.frx":24BF
      MousePointer    =   99  'Custom
      Picture         =   "FrmControles.frx":2901
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      DownPicture     =   "FrmControles.frx":3233
      Height          =   735
      Left            =   6720
      MouseIcon       =   "FrmControles.frx":4D15
      MousePointer    =   99  'Custom
      Picture         =   "FrmControles.frx":5157
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5520
      Width           =   975
   End
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   375
      Left            =   3480
      Top             =   6960
      Visible         =   0   'False
      Width           =   3375
      _ExtentX        =   5953
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
      Caption         =   "AdoDatosEmpresa"
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
   Begin MSAdodcLib.Adodc DtaControles 
      Height          =   375
      Left            =   240
      Top             =   8040
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "DtaControles"
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
   Begin MSComDlg.CommonDialog CMRutaFoto 
      Left            =   120
      Top             =   5760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      MaxFileSize     =   256
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   11880
      TabIndex        =   2
      Top             =   6600
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Generales de la Empesa"
      TabPicture(0)   =   "FrmControles.frx":5D99
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Configuracion del Sistema"
      TabPicture(1)   =   "FrmControles.frx":5DB5
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "ChkPuntos"
      Tab(1).Control(1)=   "Frame10"
      Tab(1).Control(2)=   "Frame8"
      Tab(1).Control(3)=   "Frame3"
      Tab(1).Control(4)=   "ChkTasa"
      Tab(1).Control(5)=   "Frame2"
      Tab(1).Control(6)=   "Frame1"
      Tab(1).ControlCount=   7
      TabCaption(2)   =   "Calculos"
      TabPicture(2)   =   "FrmControles.frx":5DD1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      TabCaption(3)   =   "Conexiones"
      TabPicture(3)   =   "FrmControles.frx":5DED
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame9"
      Tab(3).Control(1)=   "Skin1"
      Tab(3).ControlCount=   2
      Begin VB.Frame Frame4 
         Height          =   2775
         Left            =   -74880
         TabIndex        =   34
         Top             =   360
         Width           =   6855
         Begin VB.TextBox TxtI 
            Height          =   285
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   39
            Top             =   1800
            Width           =   3135
         End
         Begin VB.TextBox TxtC 
            Height          =   285
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   38
            Top             =   2280
            Width           =   3135
         End
         Begin VB.TextBox TxtOI 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   37
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox TxtCO 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   36
            Top             =   720
            Width           =   3135
         End
         Begin VB.TextBox TxtIO 
            Height          =   405
            Left            =   3600
            MaxLength       =   50
            TabIndex        =   35
            Top             =   240
            Width           =   3135
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":5E09
            TabIndex        =   40
            Top             =   240
            Width           =   3375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":5EB5
            TabIndex        =   41
            Top             =   720
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":5F5D
            TabIndex        =   42
            Top             =   1200
            Width           =   3495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":6013
            TabIndex        =   43
            Top             =   1800
            Width           =   3255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":60A9
            TabIndex        =   44
            Top             =   2280
            Width           =   3495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Dias al Mes"
         Height          =   1335
         Left            =   -74880
         TabIndex        =   30
         Top             =   660
         Width           =   1575
         Begin VB.OptionButton Opt30 
            Caption         =   "30 Días"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton OptExacto 
            Caption         =   "(365/12) Días"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Opt25 
            Caption         =   "25 Días"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   960
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Dias a la Semana"
         Height          =   975
         Left            =   -72960
         TabIndex        =   27
         Top             =   840
         Width           =   1695
         Begin VB.OptionButton Opt7 
            Caption         =   "Siete Dias"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton OPT6 
            Caption         =   "Seis Dias"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   600
            Width           =   1095
         End
      End
      Begin VB.CheckBox ChkTasa 
         Caption         =   "Verificar Tasa al Entrar"
         Height          =   495
         Left            =   -70200
         TabIndex        =   26
         ToolTipText     =   "Verifica si la tasa del día ya ha sido Grabada"
         Top             =   780
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Configuracion de Reportes"
         Height          =   2775
         Left            =   -74880
         TabIndex        =   14
         Top             =   2040
         Width           =   4455
         Begin VB.ComboBox CmbColillas 
            Height          =   315
            ItemData        =   "FrmControles.frx":613B
            Left            =   1560
            List            =   "FrmControles.frx":614E
            TabIndex        =   22
            Text            =   "Predeterminado"
            Top             =   360
            Width           =   2775
         End
         Begin VB.ComboBox CmbNominas 
            Height          =   315
            ItemData        =   "FrmControles.frx":61C7
            Left            =   1560
            List            =   "FrmControles.frx":61DA
            TabIndex        =   21
            Text            =   "Predeterminado"
            Top             =   840
            Width           =   2775
         End
         Begin VB.Frame Frame6 
            Height          =   1455
            Left            =   120
            TabIndex        =   15
            Top             =   1200
            Width           =   4095
            Begin VB.CommandButton Command2 
               Caption         =   "Procesar"
               Height          =   735
               Left            =   120
               Picture         =   "FrmControles.frx":624E
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   120
               Width           =   1215
            End
            Begin VB.CheckBox Chk7mo 
               Caption         =   "No Calc 7mo conforme Produccion"
               Height          =   375
               Left            =   1080
               TabIndex        =   16
               Top             =   960
               Width           =   2895
            End
            Begin MSComCtl2.DTPicker DTFecha 
               Height          =   285
               Left            =   2160
               TabIndex        =   17
               Top             =   240
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   503
               _Version        =   393216
               Format          =   78249985
               CurrentDate     =   39737
            End
            Begin XtremeSuiteControls.ProgressBar Barra 
               Height          =   255
               Left            =   1680
               TabIndex        =   19
               Top             =   600
               Visible         =   0   'False
               Width           =   2295
               _Version        =   786432
               _ExtentX        =   4048
               _ExtentY        =   450
               _StockProps     =   93
               BackColor       =   14737632
               Scrolling       =   1
               Appearance      =   6
            End
            Begin VB.Label Label3 
               Caption         =   "Hasta"
               Height          =   255
               Left            =   1560
               TabIndex        =   20
               Top             =   240
               Width           =   495
            End
         End
         Begin VB.Label Label1 
            Caption         =   "Formatos Colillas"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label2 
            Caption         =   "Formatos Nominas"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   840
            Width           =   1455
         End
         Begin VB.Label LblProcesos 
            Height          =   255
            Left            =   1680
            TabIndex        =   23
            Top             =   360
            Visible         =   0   'False
            Width           =   2415
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Configuracion Vacaciones"
         Height          =   1095
         Left            =   -70320
         TabIndex        =   11
         Top             =   1680
         Width           =   2415
         Begin VB.OptionButton OptVacacionesSemestrales 
            Caption         =   "Vacaciones Semestrales"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.OptionButton OptVacacionesMensuales 
            Caption         =   "Vacaciones Mensuales"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   2175
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "CONEXION SISTEMA CONTABLE"
         Height          =   2055
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   6855
         Begin VB.TextBox TxtConexionString 
            Height          =   1515
            Left            =   1920
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   9
            Top             =   360
            Width           =   3855
         End
         Begin VB.CommandButton Command3 
            Height          =   375
            Left            =   5880
            Picture         =   "FrmControles.frx":67D8
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   360
            Width           =   375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
            Height          =   375
            Left            =   120
            OleObjectBlob   =   "FrmControles.frx":6C8E
            TabIndex        =   10
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Configuracion del IR"
         Height          =   1095
         Left            =   -70320
         TabIndex        =   4
         Top             =   2880
         Width           =   2295
         Begin VB.OptionButton Option2 
            Caption         =   "Calcular IR x 12"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   720
            Width           =   1815
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Calcular Ajustando IR"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1935
         End
      End
      Begin VB.CheckBox ChkPuntos 
         Caption         =   "Calcular Sistema Puntos"
         Height          =   495
         Left            =   -70320
         TabIndex        =   3
         ToolTipText     =   "Verifica si la tasa del día ya ha sido Grabada"
         Top             =   4080
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   -74880
         OleObjectBlob   =   "FrmControles.frx":6D0A
         Top             =   840
      End
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4215
      Left            =   120
      TabIndex        =   47
      Top             =   1200
      Width           =   7575
      _Version        =   786432
      _ExtentX        =   13361
      _ExtentY        =   7435
      _StockProps     =   68
      Appearance      =   9
      Color           =   64
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   3
      Item(0).Caption =   "Datos Generales"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "Frame5"
      Item(1).Caption =   "Configuracion"
      Item(1).ControlCount=   5
      Item(1).Control(0)=   "GroupBox2"
      Item(1).Control(1)=   "GroupBox3"
      Item(1).Control(2)=   "GroupBox1"
      Item(1).Control(3)=   "Frame7"
      Item(1).Control(4)=   "Frame13"
      Item(2).Caption =   "Conexion"
      Item(2).ControlCount=   6
      Item(2).Control(0)=   "Label22"
      Item(2).Control(1)=   "LblMonto"
      Item(2).Control(2)=   "Frame11"
      Item(2).Control(3)=   "CmdCambiar"
      Item(2).Control(4)=   "Frame12"
      Item(2).Control(5)=   "Frame14"
      Begin VB.Frame Frame14 
         BackColor       =   &H009CCFBD&
         Caption         =   "Conexion SQL Reportes"
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   -69760
         TabIndex        =   115
         Top             =   1320
         Visible         =   0   'False
         Width           =   7215
         Begin VB.CheckBox ChkConectarSQL 
            BackColor       =   &H009CCFBD&
            Caption         =   "Conectar a SQL para Reportes"
            Height          =   255
            Left            =   240
            TabIndex        =   118
            Top             =   840
            Width           =   2535
         End
         Begin VB.CommandButton Command4 
            Caption         =   "..."
            Height          =   375
            Left            =   6480
            TabIndex        =   117
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtConexionReporte 
            Height          =   555
            Left            =   240
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   116
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H0073DBFF&
         Caption         =   "Configuracion Enlace Nominas"
         Height          =   3495
         Left            =   -66040
         TabIndex        =   112
         Top             =   480
         Visible         =   0   'False
         Width           =   3135
         Begin VB.TextBox TxtAnteponerCeros 
            Height          =   285
            Left            =   1560
            TabIndex        =   114
            Text            =   "5"
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label23 
            BackColor       =   &H0073DBFF&
            Caption         =   "Anteponer Ceros"
            Height          =   255
            Left            =   240
            TabIndex        =   113
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H009CCFBD&
         Caption         =   "Conexion Sistema de Nominas"
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   -69760
         TabIndex        =   107
         Top             =   360
         Visible         =   0   'False
         Width           =   7215
         Begin VB.CommandButton Command1 
            Caption         =   "..."
            Height          =   375
            Left            =   6480
            TabIndex        =   109
            Top             =   240
            Width           =   495
         End
         Begin VB.TextBox TxtCadena 
            Height          =   555
            Left            =   240
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   108
            Top             =   240
            Width           =   6015
         End
      End
      Begin VB.CommandButton CmdCambiar 
         Caption         =   "Cambiar"
         Height          =   375
         Left            =   -63760
         TabIndex        =   106
         Top             =   3720
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H009CCFBD&
         Caption         =   "Ubicacion de las Base de Datos"
         ForeColor       =   &H00000000&
         Height          =   975
         Left            =   -69760
         TabIndex        =   101
         Top             =   2640
         Visible         =   0   'False
         Width           =   7215
         Begin VB.CommandButton CmdDirectorio 
            Enabled         =   0   'False
            Height          =   375
            Left            =   6720
            Picture         =   "FrmControles.frx":262537
            Style           =   1  'Graphical
            TabIndex        =   105
            Top             =   480
            Width           =   375
         End
         Begin VB.TextBox TxtRuta 
            Enabled         =   0   'False
            Height          =   495
            Left            =   2160
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   104
            Top             =   360
            Width           =   4455
         End
         Begin VB.OptionButton OptRuta 
            BackColor       =   &H009CCFBD&
            Caption         =   "Ubicacion Directorio"
            Height          =   255
            Left            =   240
            TabIndex        =   103
            Top             =   600
            Width           =   1815
         End
         Begin VB.OptionButton OptRutaDefecto 
            BackColor       =   &H009CCFBD&
            Caption         =   "Ruta por Defecto"
            Height          =   255
            Left            =   240
            TabIndex        =   102
            Top             =   240
            Value           =   -1  'True
            Width           =   1815
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H0073DBFF&
         Caption         =   "Configuracion de Reportes"
         Height          =   3495
         Left            =   -69880
         TabIndex        =   96
         Top             =   480
         Visible         =   0   'False
         Width           =   3735
         Begin VB.TextBox TxtMinutosExtra 
            Height          =   285
            Left            =   1800
            TabIndex        =   110
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox TxtNoMarco 
            Height          =   285
            Left            =   1800
            TabIndex        =   99
            Top             =   720
            Width           =   975
         End
         Begin XtremeSuiteControls.CheckBox ChkRestarTolerancia 
            Height          =   255
            Left            =   240
            TabIndex        =   97
            Top             =   360
            Width           =   3255
            _Version        =   786432
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "NO Restar Minutos Tolerancia al Total"
            BackColor       =   7592959
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label20 
            BackColor       =   &H0073DBFF&
            Caption         =   "Minutos Horas Extras"
            Height          =   255
            Left            =   240
            TabIndex        =   111
            Top             =   1080
            Width           =   1575
         End
         Begin VB.Label Label19 
            BackColor       =   &H0073DBFF&
            Caption         =   "Simbolo No Marco"
            Height          =   255
            Left            =   240
            TabIndex        =   98
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Height          =   3375
         Left            =   120
         TabIndex        =   50
         Top             =   600
         Width           =   7215
         Begin VB.TextBox TxtNombreEmpresa 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   58
            Top             =   120
            Width           =   2655
         End
         Begin VB.CommandButton CmdBuscarLogo 
            Height          =   375
            Left            =   6360
            Picture         =   "FrmControles.frx":2629ED
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   2280
            Width           =   375
         End
         Begin VB.PictureBox ImgLogo2 
            AutoSize        =   -1  'True
            Height          =   2055
            Left            =   120
            ScaleHeight     =   1995
            ScaleWidth      =   2235
            TabIndex        =   56
            Top             =   240
            Width           =   2295
            Begin VB.Image ImgLogo 
               Height          =   2055
               Left            =   0
               Stretch         =   -1  'True
               Top             =   0
               Width           =   2295
            End
         End
         Begin VB.TextBox TxtRutaLogo 
            Height          =   375
            Left            =   2880
            TabIndex        =   55
            Top             =   2280
            Width           =   3495
         End
         Begin VB.TextBox TxtFax 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   54
            Top             =   1560
            Width           =   2655
         End
         Begin VB.TextBox TxtDireccionEmpresa 
            Height          =   765
            Left            =   4080
            MaxLength       =   50
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   53
            Top             =   360
            Width           =   2655
         End
         Begin VB.TextBox TxtRucEmpresa 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   52
            Top             =   1080
            Width           =   2655
         End
         Begin VB.TextBox TxtTelefono 
            Height          =   285
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   51
            Top             =   1320
            Width           =   2655
         End
         Begin XtremeSuiteControls.CheckBox ChkMembreteLogo 
            Height          =   255
            Left            =   2880
            TabIndex        =   100
            Top             =   2880
            Width           =   3255
            _Version        =   786432
            _ExtentX        =   5741
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Utilizar logo como Membrete de Reportes"
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label9 
            Caption         =   "Dirección de Logo de la Empresa"
            Height          =   255
            Left            =   240
            TabIndex        =   64
            Top             =   2400
            Width           =   2535
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Fax"
            Height          =   255
            Left            =   3600
            TabIndex        =   63
            Top             =   1560
            Width           =   375
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Telefono"
            Height          =   255
            Left            =   3240
            TabIndex        =   62
            Top             =   1320
            Width           =   735
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Numero RUC"
            Height          =   255
            Left            =   3000
            TabIndex        =   61
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Direccion"
            Height          =   255
            Left            =   3240
            TabIndex        =   60
            Top             =   360
            Width           =   735
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre Empresa"
            Height          =   255
            Left            =   2760
            TabIndex        =   59
            Top             =   120
            Width           =   1215
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox2 
         Height          =   3495
         Left            =   -65320
         TabIndex        =   65
         Top             =   4560
         Visible         =   0   'False
         Width           =   2895
         _Version        =   786432
         _ExtentX        =   5106
         _ExtentY        =   6165
         _StockProps     =   79
         Caption         =   "Configuracion Horas Almuerzo"
         BackColor       =   7592959
         Begin MSMask.MaskEdBox TxtSalida 
            Height          =   285
            Left            =   1680
            TabIndex        =   66
            Top             =   840
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtEntrada 
            Height          =   285
            Left            =   1680
            TabIndex        =   69
            Top             =   480
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtEntrada2 
            Height          =   285
            Left            =   1680
            TabIndex        =   70
            Top             =   2040
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtEntrada1 
            Height          =   285
            Left            =   1680
            TabIndex        =   73
            Top             =   1680
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtSalida2 
            Height          =   285
            Left            =   1680
            TabIndex        =   74
            Top             =   3000
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox TxtSalida1 
            Height          =   285
            Left            =   1680
            TabIndex        =   77
            Top             =   2640
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   503
            _Version        =   393216
            MaxLength       =   5
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ventana Entrada"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   79
            Top             =   1440
            Width           =   2295
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Ventana Salida"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   240
            TabIndex        =   78
            Top             =   2400
            Width           =   2295
         End
         Begin VB.Label Label16 
            BackStyle       =   0  'Transparent
            Caption         =   "Fuera Serv Entrada"
            Height          =   375
            Left            =   120
            TabIndex        =   76
            Top             =   3000
            Width           =   1455
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Servicio Entrada"
            Height          =   375
            Left            =   120
            TabIndex        =   75
            Top             =   2640
            Width           =   1335
         End
         Begin VB.Label Label14 
            BackStyle       =   0  'Transparent
            Caption         =   "Fuera Serv Entrada"
            Height          =   375
            Left            =   240
            TabIndex        =   72
            Top             =   2040
            Width           =   1455
         End
         Begin VB.Label Label13 
            BackStyle       =   0  'Transparent
            Caption         =   "Servicio Entrada"
            Height          =   375
            Left            =   240
            TabIndex        =   71
            Top             =   1680
            Width           =   1335
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrada Almuerzo"
            Height          =   375
            Left            =   240
            TabIndex        =   68
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Salida Almuerzo"
            Height          =   375
            Left            =   240
            TabIndex        =   67
            Top             =   840
            Width           =   1215
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox3 
         Height          =   1935
         Left            =   -69760
         TabIndex        =   80
         Top             =   4440
         Visible         =   0   'False
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Horarios Asignados - Almuerzos"
         BackColor       =   7592959
         Begin XtremeSuiteControls.PushButton CmdAgregar 
            Height          =   375
            Left            =   240
            TabIndex        =   84
            Top             =   1320
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Grabar"
            UseVisualStyle  =   -1  'True
         End
         Begin TrueOleDBList80.TDBCombo TDBMatutino 
            Bindings        =   "FrmControles.frx":262EA3
            Height          =   315
            Left            =   1680
            TabIndex        =   83
            Top             =   360
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   0
            _EDITHEIGHT     =   556
            _GAPHEIGHT      =   53
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).DataField=   ""
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   0
            Splits(0).ExtendRightColumn=   -1  'True
            Splits(0).AllowRowSizing=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
            Splits.Count    =   1
            Appearance      =   1
            BorderStyle     =   1
            ComboStyle      =   2
            AutoCompletion  =   0   'False
            LimitToList     =   0   'False
            ColumnHeaders   =   -1  'True
            ColumnFooters   =   0   'False
            DataMode        =   0
            DefColWidth     =   0
            Enabled         =   -1  'True
            HeadLines       =   1
            FootLines       =   1
            RowDividerStyle =   0
            Caption         =   ""
            EditFont        =   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
            LayoutName      =   ""
            LayoutFileName  =   ""
            MultipleLines   =   0
            EmptyRows       =   -1  'True
            CellTips        =   0
            AutoSize        =   -1  'True
            ListField       =   "Schname"
            BoundColumn     =   ""
            IntegralHeight  =   0   'False
            CellTipsWidth   =   0
            CellTipsDelay   =   1000
            AutoDropdown    =   0   'False
            RowTracking     =   -1  'True
            RightToLeft     =   0   'False
            RowMember       =   ""
            MouseIcon       =   0
            MouseIcon.vt    =   3
            MousePointer    =   0
            MatchEntryTimeout=   2000
            OLEDragMode     =   0
            OLEDropMode     =   0
            AnimateWindow   =   0
            AnimateWindowDirection=   0
            AnimateWindowTime=   200
            AnimateWindowClose=   0
            DropdownPosition=   0
            Locked          =   0   'False
            ScrollTrack     =   0   'False
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            AddItemSeparator=   ";"
            _PropDict       =   $"FrmControles.frx":262EBD
            _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
            _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
            _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
            _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
            _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0"
            _StyleDefs(40)  =   "Named:id=34:Heading"
            _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(42)  =   ":id=34,.wraptext=-1"
            _StyleDefs(43)  =   "Named:id=35:Footing"
            _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(45)  =   "Named:id=36:Selected"
            _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(47)  =   "Named:id=37:Caption"
            _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(49)  =   "Named:id=38:HighlightRow"
            _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(51)  =   "Named:id=39:EvenRow"
            _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(53)  =   "Named:id=40:OddRow"
            _StyleDefs(54)  =   ":id=40,.parent=33"
            _StyleDefs(55)  =   "Named:id=41:RecordSelector"
            _StyleDefs(56)  =   ":id=41,.parent=34"
            _StyleDefs(57)  =   "Named:id=42:FilterBar"
            _StyleDefs(58)  =   ":id=42,.parent=33"
         End
         Begin XtremeSuiteControls.CheckBox ChkUnicoHorario 
            Height          =   255
            Left            =   360
            TabIndex        =   82
            Top             =   720
            Width           =   3615
            _Version        =   786432
            _ExtentX        =   6376
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Excluir Sabados del Horario Almuerzo"
            BackColor       =   7592959
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton CmdEliminar 
            Height          =   375
            Left            =   2880
            TabIndex        =   85
            Top             =   1320
            Width           =   1215
            _Version        =   786432
            _ExtentX        =   2143
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Eliminar"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.CheckBox ChkAsignarHorario 
            Height          =   255
            Left            =   360
            TabIndex        =   86
            Top             =   1020
            Width           =   3615
            _Version        =   786432
            _ExtentX        =   6376
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Asignar Almuerzo al Personal sin Horario"
            BackColor       =   7592959
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label LblMatutino 
            BackStyle       =   0  'Transparent
            Caption         =   "Horario Matutino:"
            Height          =   375
            Left            =   240
            TabIndex        =   81
            Top             =   360
            Width           =   1335
         End
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   1935
         Left            =   -69760
         TabIndex        =   87
         Top             =   4320
         Visible         =   0   'False
         Width           =   4215
         _Version        =   786432
         _ExtentX        =   7435
         _ExtentY        =   3413
         _StockProps     =   79
         Caption         =   "Configuracion Calculo de Horas Extras"
         BackColor       =   7592959
         Begin VB.TextBox TxtHorasDomingo 
            Height          =   285
            Left            =   1440
            TabIndex        =   94
            Text            =   "0"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtHorasSabado 
            Height          =   285
            Left            =   1440
            TabIndex        =   93
            Text            =   "3"
            Top             =   960
            Width           =   735
         End
         Begin VB.TextBox TxtHorasTrab 
            Height          =   285
            Left            =   1440
            TabIndex        =   88
            Text            =   "8"
            Top             =   360
            Width           =   735
         End
         Begin XtremeSuiteControls.RadioButton OptHoras 
            Height          =   495
            Left            =   2280
            TabIndex        =   89
            Top             =   240
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Calcular Por Horas Trabajadas"
            BackColor       =   7592959
            UseVisualStyle  =   -1  'True
            Value           =   -1  'True
         End
         Begin XtremeSuiteControls.RadioButton OptHoraSalida 
            Height          =   495
            Left            =   2280
            TabIndex        =   90
            Top             =   840
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Calcular Por Hora Salida"
            BackColor       =   7592959
            UseVisualStyle  =   -1  'True
         End
         Begin VB.Label Label11 
            BackStyle       =   0  'Transparent
            Caption         =   "Calcular Horas Extras Domingo"
            Height          =   375
            Left            =   240
            TabIndex        =   95
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label LblHora 
            BackStyle       =   0  'Transparent
            Caption         =   "Calcular Horas Extras Sabado"
            Height          =   375
            Left            =   240
            TabIndex        =   92
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label LblHorasTrab 
            BackStyle       =   0  'Transparent
            Caption         =   "Calcular Horas Extras despues"
            Height          =   375
            Left            =   240
            TabIndex        =   91
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Label LblMonto 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   -60520
         TabIndex        =   49
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Monto Total Acumulado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   -63640
         TabIndex        =   48
         Top             =   5040
         Visible         =   0   'False
         Width           =   3135
      End
   End
   Begin vbskfree.Skinner Skinner1 
      Left            =   240
      Top             =   5040
      _ExtentX        =   1270
      _ExtentY        =   1270
      CloseButtonToolTipText=   "Cerrar"
      MinButtonToolTipText=   "Minimizar"
      MaxButtonToolTipText=   "Maximizar"
      RestoreButtonToolTipText=   "Restaurar"
   End
   Begin MSAdodcLib.Adodc AdoHorarios 
      Height          =   375
      Left            =   3120
      Top             =   8520
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoHorarios"
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
   Begin MSAdodcLib.Adodc AdoHorarios2 
      Height          =   375
      Left            =   3360
      Top             =   7920
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
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
      Caption         =   "AdoHorarios2"
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   480
      Top             =   7320
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoConsulta"
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
   Begin MSAdodcLib.Adodc AdoConexion 
      Height          =   375
      Left            =   3120
      Top             =   7440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
      Caption         =   "AdoConexion"
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
Attribute VB_Name = "FrmControlesReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ChkConectarSQL_Click()
  If Me.ChkConectarSQL.Value = xtpUnchecked Then
  
  
  End If

End Sub

Private Sub CmbNominas_Change()
  If Me.CmbNominas.Text = "Nomina Bono Produccion" Then
    Me.Frame6.Visible = True
  Else
    Me.Frame6.Visible = False
  End If
End Sub

Private Sub CmbNominas_Click()
  If Me.CmbNominas.Text = "Nomina Bono Produccion" Then
    Me.Frame6.Visible = True
  Else
    Me.Frame6.Visible = False
  End If
End Sub

Private Sub CmdAceptar_Click()
'On Error GoTo TipoErr

'DtaControles.Recordset.Edit


Me.AdoDatosEmpresa.Refresh
If Not Me.AdoDatosEmpresa.Recordset.EOF Then
 Me.AdoDatosEmpresa.Recordset("NombreEmpresa") = Me.TxtNombreEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("RutaLogo") = Me.TxtRutaLogo.Text
 Me.AdoDatosEmpresa.Recordset("NumeroRUC") = Me.TxtRucEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("Direccion") = Me.TxtDireccionEmpresa.Text
 If Me.TxtTelefono.Text <> "" Then
 Me.AdoDatosEmpresa.Recordset("Telefono") = Me.TxtTelefono.Text
 End If
 If Me.TxtAnteponerCeros.Text <> "" Then
  Me.AdoDatosEmpresa.Recordset("AnteponerCeros") = Me.TxtAnteponerCeros.Text
 Else
  Me.AdoDatosEmpresa.Recordset("AnteponerCeros") = 5
 End If
 
 If Me.TxtFax.Text <> "" Then
  Me.AdoDatosEmpresa.Recordset("Fax") = Me.TxtFax.Text
 End If
 
  If Me.OptHoras.Value = False Then
    Me.AdoDatosEmpresa.Recordset("CalcularHorasTrab") = 0
  Else
    Me.AdoDatosEmpresa.Recordset("CalcularHorasTrab") = 1
 End If
 
 If Me.ChkRestarTolerancia.Value = xtpUnchecked Then
    Me.AdoDatosEmpresa.Recordset("RestarToleranciaLlegada") = 0
  Else
    Me.AdoDatosEmpresa.Recordset("RestarToleranciaLlegada") = 1
 End If
 
 If Me.ChkMembreteLogo.Value = xtpUnchecked Then
    Me.AdoDatosEmpresa.Recordset("MembreteLogo") = 0
  Else
    Me.AdoDatosEmpresa.Recordset("MembreteLogo") = 1
 End If
 
  If Me.ChkConectarSQL.Value = xtpUnchecked Then
    Me.AdoDatosEmpresa.Recordset("ConexioSQL") = 0
  Else
    Me.AdoDatosEmpresa.Recordset("ConexioSQL") = 1
 End If
 
 
 If Me.TxtHorasTrab.Text <> "" Then
  Me.AdoDatosEmpresa.Recordset("MinutosExtra") = Me.TxtMinutosExtra.Text
 End If
 
 If Me.TxtHorasTrab.Text <> "" Then
  Me.AdoDatosEmpresa.Recordset("HorasTrab") = Me.TxtHorasTrab.Text
 End If
 
 If Me.TxtHorasSabado.Text <> "" Then
  Me.AdoDatosEmpresa.Recordset("HorasTrabSab") = Me.TxtHorasSabado.Text
 End If
 
 If Me.TxtHorasDomingo.Text <> "" Then
  Me.AdoDatosEmpresa.Recordset("HorasTrabDom") = Me.TxtHorasDomingo.Text
 End If
 
 If Me.TxtNoMarco.Text <> "" Then
  Me.AdoDatosEmpresa.Recordset("SimboloNoMarco") = Me.TxtNoMarco.Text
 End If
' If Me.TxtHora.Text <> "" Then
'  Me.AdoDatosEmpresa.Recordset("HoraSalida") = Me.TxtHora.Text
' End If

 If Me.TxtCadena.Text <> "" Then
   Me.AdoDatosEmpresa.Recordset("Cadena") = Me.TxtCadena.Text
 End If
 
  If Me.TxtCadena.Text <> "" Then
   Me.AdoDatosEmpresa.Recordset("Cadena2") = Me.TxtConexionReporte.Text
 End If
 
 Me.AdoDatosEmpresa.Recordset.Update


Else
 Me.AdoDatosEmpresa.Recordset.AddNew
 Me.AdoDatosEmpresa.Recordset("Numero") = 1
 Me.AdoDatosEmpresa.Recordset("NombreEmpresa") = Me.TxtNombreEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("RutaLogo") = Me.TxtRutaLogo.Text
 Me.AdoDatosEmpresa.Recordset("NumeroRUC") = Me.TxtRucEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("Direccion") = Me.TxtDireccionEmpresa.Text
 Me.AdoDatosEmpresa.Recordset("Telefono") = Me.TxtTelefono.Text
 Me.AdoDatosEmpresa.Recordset("Fax") = Me.TxtFax.Text
 Me.AdoDatosEmpresa.Recordset.Update
End If






RutaLogo = Me.TxtRutaLogo.Text


Unload Me

Exit Sub
TipoErr:
    MsgBox Err.Description
End Sub

Private Sub CmdAgregar_Click()
  Dim Codigo As Double
  Codigo = Me.TDBMatutino.Columns(0).Text
  Me.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & Codigo & "))"
  Me.AdoConsulta.Refresh
  If Me.AdoConsulta.Recordset.EOF Then
        Me.AdoHorarios2.Refresh
        Me.AdoHorarios2.Recordset.AddNew
        Me.AdoHorarios2.Recordset("Schid") = Me.TDBMatutino.Columns(0).Text
        Me.AdoHorarios2.Recordset("EntradaAlmuerzo") = Me.TxtEntrada.Text
        Me.AdoHorarios2.Recordset("SalidaAlmuerzo") = Me.TxtSalida.Text
        Me.AdoHorarios2.Recordset("EntradaAlmuerzo1") = Me.TxtEntrada1.Text
        Me.AdoHorarios2.Recordset("EntradaAlmuerzo2") = Me.TxtEntrada2.Text
        Me.AdoHorarios2.Recordset("SalidaAlmuerzo1") = Me.TxtSalida1.Text
        Me.AdoHorarios2.Recordset("SalidaAlmuerzo2") = Me.TxtSalida2.Text
        If Me.ChkUnicoHorario.Value = xtpChecked Then
         Me.AdoHorarios2.Recordset("ExcluirSabado") = True
        Else
         Me.AdoHorarios2.Recordset("ExcluirSabado") = False
        End If
        
        If Me.ChkAsignarHorario.Value = xtpChecked Then
          Me.AdoHorarios2.Recordset("PersonalSinHorario") = True
        Else
          Me.AdoHorarios2.Recordset("PersonalSinHorario") = False
        End If
        
        Me.AdoHorarios2.Recordset.Update
   Else
        Me.AdoHorarios2.Recordset("EntradaAlmuerzo") = Me.TxtEntrada.Text
        Me.AdoHorarios2.Recordset("SalidaAlmuerzo") = Me.TxtSalida.Text
        Me.AdoHorarios2.Recordset("EntradaAlmuerzo1") = Me.TxtEntrada1.Text
        Me.AdoHorarios2.Recordset("EntradaAlmuerzo2") = Me.TxtEntrada2.Text
        Me.AdoHorarios2.Recordset("SalidaAlmuerzo1") = Me.TxtSalida1.Text
        Me.AdoHorarios2.Recordset("SalidaAlmuerzo2") = Me.TxtSalida2.Text
        If Me.ChkUnicoHorario.Value = xtpChecked Then
         Me.AdoHorarios2.Recordset("ExcluirSabado") = True
        Else
         Me.AdoHorarios2.Recordset("ExcluirSabado") = False
        End If
        
        If Me.ChkAsignarHorario.Value = xtpChecked Then
          Me.AdoHorarios2.Recordset("PersonalSinHorario") = True
        Else
          Me.AdoHorarios2.Recordset("PersonalSinHorario") = False
        End If
        Me.AdoHorarios2.Recordset.Update
   End If

MsgBox "Registro Grabado", vbExclamation, "Reportes Zeus"
End Sub

Private Sub CmdBuscarLogo_Click()
Dim retval
Dim OpenFileName As String
    On Error Resume Next
    ' Set the commom dialog properties we need
    If Me.TxtRutaLogo.Text <> "" Then
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text
    End If
    CMRutaFoto.FileName = ""
    ' We will load BMP, JPG, and TIF files
    
    CMRutaFoto.Filter = "Image Files |*.bmp;*.gif;*.jpg;*.png;*.tif|All files |*.*"
    ' Display common dialog box
    CMRutaFoto.ShowOpen
    Me.TxtRutaLogo.Text = CMRutaFoto.FileName
    
    Me.ImgLogo.Picture = LoadPicture(Me.TxtRutaLogo.Text)
End Sub

Private Sub CmdCambiar_Click()
Dim RutaConexion As String, ConexionBD As String
Dim RutaBD As String

MDIPrimero.PopupControl1.RemoveAllItems
RutaConexion = App.Path + "\CntReloj.dll"
ConexionBD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaConexion & ";Persist Security Info=False"
With Me.AdoConexion
 .ConnectionString = ConexionBD
 .RecordSource = "SELECT Servidor.* FROM Servidor"
 .Refresh
End With


  If Me.OptRutaDefecto.Value = True Then
    RutaBD = "APP"
  Else
    If Me.TxtRuta.Text = "" Then
     MsgBox "Necesita Indicar la Ruta", vbCritical, "Zeus Reloj"
     Exit Sub
    Else
     RutaBD = Me.TxtRuta.Text
    End If
  End If
  
  
  If Me.AdoConexion.Recordset.EOF Then
    Me.AdoConexion.Recordset.AddNew
    Me.AdoConexion.Recordset("Servidor") = RutaBD
    Me.AdoConexion.Recordset.Update
  Else
    Me.AdoConexion.Recordset("Servidor") = RutaBD
    Me.AdoConexion.Recordset.Update
  End If
  
  '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
  '////////////////////////////////////////////CAMBIO LA CONFIGURACION DEL SISTEMA ////////////////////////////////////////
  '////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

         
          If RutaBD = "APP" Then
            RutaServerReloj = App.Path + "\Att2007.mdb"
            RutaServerEasy = App.Path + "\Att2003.mdb"
            RutaBD = App.Path
          Else
            RutaServerReloj = RutaBD + "\Att2007.mdb"
            RutaServerEasy = RutaBD + "\Att2003.mdb"
          End If
          
          ConexionEasy = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaServerEasy & ";Persist Security Info=False"
          ConexionReloj = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaServerReloj & ";Persist Security Info=False"
          
          With MDIPrimero.DtaEmpresa
           .ConnectionString = ConexionReloj
           .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
           .Refresh
          End With
          
          With MDIPrimero.AdoConsultaEasyWay
             .ConnectionString = ConexionEasy
          End With
          
          
          With MDIPrimero.AdoDispositivos
             .ConnectionString = ConexionEasy
          End With
        
          '//////////////////////////////////BUSCO EL NOMBRE DE LA EMPRESA ///////////////////////////
          MDIPrimero.AdoConsultaEasyWay.RecordSource = "SELECT Dept.DeptName, Dept.SupDeptid From Dept WHERE (((Dept.SupDeptid)=0))"
          MDIPrimero.AdoConsultaEasyWay.Refresh
          If Not MDIPrimero.AdoConsultaEasyWay.Recordset.EOF Then
            MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa") = MDIPrimero.AdoConsultaEasyWay.Recordset("DeptName")
            MDIPrimero.DtaEmpresa.Recordset.Update
          End If
          
          '//////////////////////////////////BUSCO SI EXISTE EL SISTEMA PARA AGREGARLO AL MENU DEL FABRICANTE ///////////////////////////
          MDIPrimero.AdoConsultaEasyWay.RecordSource = "SELECT OutProg.Progid, OutProg.ProgName, OutProg.ProgPath From OutProg WHERE (((OutProg.ProgName)='REPORTES ZEUS RELOJ')) "
          MDIPrimero.AdoConsultaEasyWay.Refresh
          If MDIPrimero.AdoConsultaEasyWay.Recordset.EOF Then
            MDIPrimero.AdoConsultaEasyWay.Recordset.AddNew
              MDIPrimero.AdoConsultaEasyWay.Recordset("ProgName") = "REPORTES ZEUS RELOJ"
              MDIPrimero.AdoConsultaEasyWay.Recordset("ProgPath") = RutaBD + "\Zeus Reloj.exe"
            MDIPrimero.AdoConsultaEasyWay.Recordset.Update
          Else
            MDIPrimero.AdoConsultaEasyWay.Recordset("ProgPath") = RutaBD + "\Zeus Reloj.exe"
            MDIPrimero.AdoConsultaEasyWay.Recordset.Update
          End If
          
          MDIPrimero.DtaEmpresa.Refresh
          Titulo = MDIPrimero.DtaEmpresa.Recordset("NombreEmpresa")
          SubTitulo = MDIPrimero.DtaEmpresa.Recordset("Direccion") + " RUC: " + MDIPrimero.DtaEmpresa.Recordset("NumeroRuc")
          ''RutaLogo = DtaEmpresa.Recordset.RutaLogo
          'StatusBar2.Panels(2) = "Licencia: " + Titulo
          
          
          Set item = MDIPrimero.PopupControl1.AddItem(50, 15, 270, 45, Titulo)
          item.TextColor = RGB(0, 61, 178)
          item.Bold = True
          
          Set item = MDIPrimero.PopupControl1.AddItem(12, 20, 12, 27, "")
'          item.SetIcon LoadIcon("Imagenes\Imagen.ico", 32, 32), xtpPopupItemIconNormal
          
          
          Set item = MDIPrimero.PopupControl1.AddItem(50, 29, 400, 100, "Direc:" & MDIPrimero.DtaEmpresa.Recordset("Direccion"))
          item.TextColor = RGB(0, 61, 178)
          item.Bold = True
          Set item = MDIPrimero.PopupControl1.AddItem(60, 60, 400, 100, "ZEUS RELOJ  ")
              item.Bold = True
              MDIPrimero.PopupControl1.VisualTheme = xtpPopupThemeOffice2003
              MDIPrimero.PopupControl1.SetSize 300, 110
              MDIPrimero.PopupControl1.Show
              MDIPrimero.PopupControl1.Show

End Sub

Private Sub CmdCancelar_Click()
Unload Me
End Sub


Private Sub Command1_Click()
On Error GoTo TipoErrs
Dim mydlg As New MSDASC.DataLinks
Dim ADOcon As New ADODB.Connection

Me.TxtCadena.Text = mydlg.PromptNew


Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub CmdDirectorio_Click()
    Dim ret As String
    ' Le pasa la leyenda del cuadro de iálogo y el path inicial
    ret = Buscar_Carpeta(" ... Seleccione una carpeta ")
  
    Me.TxtRuta.Text = ret
End Sub

Private Sub CmdEliminar_Click()
  Dim Codigo As Double, Respuesta As Integer
  Codigo = Me.TDBMatutino.Columns(0).Text
  Me.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & Codigo & "))"
  Me.AdoConsulta.Refresh
  If Not Me.AdoConsulta.Recordset.EOF Then
    Respuesta = MsgBox("Esta Seguro de Eliminar el Registro", vbYesNo, "Zeus Contable")
    If Respuesta = 6 Then
      Me.AdoConsulta.Recordset.Delete
    End If
  End If
End Sub

Private Sub Command2_Click()
 Dim Fecha As String, i As Double, HorasExtra As Double, Maximo As Double
  Me.SSTab1.Enabled = False
  
   Fecha = Format(Me.DTFecha.Value, "yyyy-mm-dd")
   Me.AdoDetalleNominas.RecordSource = "SELECT   Nomina.FechaNominaINI, Nomina.FechaNomina, DetalleNomina.Ajuste, DetalleNomina.HorasExtras FROM  DetalleNomina INNER JOIN Nomina ON DetalleNomina.NumNomina = Nomina.NumNomina WHERE (Nomina.FechaNomina <= CONVERT(DATETIME, '" & Fecha & "', 102))"
   Me.AdoDetalleNominas.Refresh
   
   Maximo = Me.AdoDetalleNominas.Recordset.RecordCount
   Me.Barra.Min = 0
   Me.Barra.Max = Maximo
   Me.Barra.Value = 0
   Me.Barra.Visible = True
   i = 0
   Me.LblProcesos.Visible = True
   
   Do While Not Me.AdoDetalleNominas.Recordset.EOF
    
    
     HorasExtra = Me.AdoDetalleNominas.Recordset("HorasExtras")
     
     Me.AdoDetalleNominas.Recordset("Ajuste") = HorasExtra
     Me.AdoDetalleNominas.Recordset.Update
     
   
     i = i + 1
     DoEvents
     Me.Barra.Value = i
     Me.LblProcesos.Caption = "Procesando " & i & " de " & Maximo
     Me.AdoDetalleNominas.Recordset.MoveNext
   Loop


  Me.SSTab1.Enabled = True
End Sub

Private Sub Command3_Click()
On Error GoTo TipoErrs
Dim mydlg As New MSDASC.DataLinks
Dim ADOcon As New ADODB.Connection

Me.TxtConexionString.Text = mydlg.PromptNew


Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub Command4_Click()
On Error GoTo TipoErrs
Dim mydlg As New MSDASC.DataLinks
Dim ADOcon As New ADODB.Connection

Me.TxtConexionReporte.Text = mydlg.PromptNew


Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo TipoErrs
Dim Destino As String
Dim RutaConexion As String, ConexionBD As String
Dim RutaBD As String
Me.Top = 1500
Me.Left = 4500


Me.TxtEntrada.Text = "00:00"
Me.TxtSalida.Text = "00:00"
Me.TxtEntrada1.Text = "00:00"
Me.TxtEntrada2.Text = "00:00"
Me.TxtSalida1.Text = "00:00"
Me.TxtSalida2.Text = "00:00"


RutaConexion = App.Path + "\CntReloj.dll"
ConexionBD = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & RutaConexion & ";Persist Security Info=False"
With Me.AdoConexion
 .ConnectionString = ConexionBD
 .RecordSource = "SELECT Servidor.* FROM Servidor"
 .Refresh
End With

If Not Me.AdoConexion.Recordset.EOF Then
  RutaBD = Me.AdoConexion.Recordset("Servidor")
  If RutaBD = "APP" Then
    Me.OptRutaDefecto.Value = True
    Me.TxtRuta.Enabled = False
    Me.CmdDirectorio.Enabled = False
  Else
    Me.OptRuta.Value = True
    Me.TxtRuta.Enabled = True
    Me.CmdDirectorio.Enabled = True
    Me.TxtRuta.Text = Me.AdoConexion.Recordset("Servidor")
  End If
End If




Me.DTFecha.Value = Format(Now, "dd/mm/yyyy")


With Me.AdoDatosEmpresa
   .ConnectionString = ConexionReloj
   .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
   .Refresh
End With

With Me.AdoHorarios
  .ConnectionString = ConexionEasy
  .RecordSource = "SELECT Schedule.Schid, Schedule.Schname FROM Schedule"
  .Refresh
End With

With Me.AdoHorarios2
  .ConnectionString = ConexionReloj
  .RecordSource = "SELECT Horario.* FROM Horario"
  .Refresh
End With

With Me.AdoConsulta
   .ConnectionString = ConexionReloj
End With


If Not Me.AdoDatosEmpresa.Recordset.EOF Then

 Me.TxtNombreEmpresa.Text = Me.AdoDatosEmpresa.Recordset("NombreEmpresa")
 Me.TxtRucEmpresa.Text = Me.AdoDatosEmpresa.Recordset("NumeroRUC")
 Me.TxtDireccionEmpresa.Text = Me.AdoDatosEmpresa.Recordset("Direccion")
 Me.TxtTelefono.Text = Me.AdoDatosEmpresa.Recordset("Telefono")
 Me.TxtFax.Text = Me.AdoDatosEmpresa.Recordset("Fax")
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("AnteponerCeros")) Then
  Me.TxtAnteponerCeros.Text = Me.AdoDatosEmpresa.Recordset("AnteponerCeros")
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("RutaLogo")) Then
  Me.TxtRutaLogo.Text = Me.AdoDatosEmpresa.Recordset("RutaLogo")
 End If
 
 If Me.AdoDatosEmpresa.Recordset("CalcularHorasTrab") = 0 Then
   Me.OptHoras.Value = False
   Me.OptHoraSalida.Value = True
 Else
   Me.OptHoras.Value = True
   Me.OptHoraSalida.Value = False
 End If
 
  If Me.AdoDatosEmpresa.Recordset("RestarToleranciaLlegada") = 0 Then
   Me.ChkRestarTolerancia.Value = xtpUnchecked
 Else
   Me.ChkRestarTolerancia.Value = xtpChecked
 End If
 
 If Me.AdoDatosEmpresa.Recordset("MembreteLogo") = 0 Then
   Me.ChkMembreteLogo.Value = xtpUnchecked
 Else
   Me.ChkMembreteLogo.Value = xtpChecked
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("MinutosExtra")) Then
  Me.TxtMinutosExtra.Text = Me.AdoDatosEmpresa.Recordset("MinutosExtra")
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("HorasTrab")) Then
  Me.TxtHorasTrab.Text = Me.AdoDatosEmpresa.Recordset("HorasTrab")
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("SimboloNoMarco")) Then
  Me.TxtNoMarco.Text = Me.AdoDatosEmpresa.Recordset("SimboloNoMarco")
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("HorasTrabDom")) Then
  Me.TxtHorasDomingo.Text = Me.AdoDatosEmpresa.Recordset("HorasTrabDom")
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("Cadena")) Then
  Me.TxtCadena.Text = Me.AdoDatosEmpresa.Recordset("Cadena")
 End If
 
 If Not IsNull(Me.AdoDatosEmpresa.Recordset("Cadena2")) Then
  Me.TxtConexionReporte.Text = Me.AdoDatosEmpresa.Recordset("Cadena2")
 End If
 
If Me.AdoDatosEmpresa.Recordset("ConexioSQL") = 0 Then
   Me.ChkConectarSQL.Value = xtpUnchecked
 Else
   Me.ChkConectarSQL.Value = xtpChecked
 End If
 
' If Not IsNull(Me.AdoDatosEmpresa.Recordset("HoraSalida")) Then
'  Me.TxtHora.Text = Format(Me.AdoDatosEmpresa.Recordset("HoraSalida"), "hh:mm")
' End If
'
    If Me.TxtRutaLogo.Text <> "" Then
        If (Dir(Me.TxtRutaLogo.Text) <> "") Then
          Me.ImgLogo.Picture = LoadPicture(Me.TxtRutaLogo.Text)
        Else
'          Destino = RutaFoto + "Zw.bmp"
'          Me.ImgLogo.Picture = LoadPicture(Destino)
          MsgBox "La Ruta del LOGO: " & Me.TxtRutaLogo & " ES INCORRECTA"
          
        End If
       CMRutaFoto.InitDir = Me.TxtRutaLogo.Text

    End If
        

    
End If





Exit Sub
TipoErrs:
MsgBox Err.Description
End Sub






Private Sub RadioButton1_Click()

End Sub

Private Sub OptHoras_Click()
  If Me.OptHoras.Value = True Then
    Me.LblHora.Enabled = True
    Me.LblHorasTrab.Enabled = True
    Me.TxtHorasTrab.Enabled = True
    Me.TxtHorasSabado.Enabled = True
  Else
     Me.LblHora.Enabled = False
    Me.LblHorasTrab.Enabled = False
    Me.TxtHorasTrab.Enabled = False
    Me.TxtHorasSabado.Enabled = False
  End If
End Sub

Private Sub OptHoraSalida_Click()
 If Me.OptHoraSalida.Value = True Then
    Me.LblHora.Enabled = False
    Me.LblHorasTrab.Enabled = False
    Me.TxtHorasTrab.Enabled = False
    Me.TxtHorasSabado.Enabled = False
 Else
    Me.LblHora.Enabled = True
    Me.LblHorasTrab.Enabled = True
    Me.TxtHorasTrab.Enabled = True
    Me.TxtHorasSabado.Enabled = True
 End If
End Sub

Private Sub OptRuta_Click()
 If Me.OptRuta.Value = True Then
   Me.TxtRuta.Enabled = True
   Me.CmdDirectorio.Enabled = True
 End If
End Sub

Private Sub OptRutaDefecto_Click()
 If Me.OptRutaDefecto.Value = True Then
   Me.TxtRuta.Enabled = False
   Me.CmdDirectorio.Enabled = False
   Me.TxtRuta.Text = ""
 End If
End Sub

Private Sub TDBMatutino_Change()
  Dim Codigo As Double
  Codigo = Me.TDBMatutino.Columns(0).Text
  Me.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & Codigo & "))"
  Me.AdoConsulta.Refresh
  If Me.AdoConsulta.Recordset.EOF Then

        Me.TxtEntrada.Text = "00:00"
        Me.TxtSalida.Text = "00:00"
        Me.TxtEntrada1.Text = "00:00"
        Me.TxtEntrada2.Text = "00:00"
        Me.TxtSalida1.Text = "00:00"
        Me.TxtSalida2.Text = "00:00"
        Me.ChkUnicoHorario.Value = xtpUnchecked
  Else
'        Me.TxtHora.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo")
        Me.TxtEntrada.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo")
        Me.TxtSalida.Text = Me.AdoConsulta.Recordset("SalidaAlmuerzo")
        Me.TxtEntrada1.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo1")
        Me.TxtEntrada2.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo2")
        Me.TxtSalida1.Text = Me.AdoConsulta.Recordset("SalidaAlmuerzo1")
        Me.TxtSalida2.Text = Me.AdoConsulta.Recordset("SalidaAlmuerzo2")
        If Me.AdoHorarios2.Recordset("ExcluirSabado") = True Then
          Me.ChkUnicoHorario.Value = xtpChecked
        Else
          Me.ChkUnicoHorario.Value = xtpUnchecked
        End If
  End If
  

End Sub

Private Sub TDBMatutino_ItemChange()
  Dim Codigo As Double
  Codigo = Me.TDBMatutino.Columns(0).Text
  Me.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & Codigo & "))"
  Me.AdoConsulta.Refresh
  If Me.AdoConsulta.Recordset.EOF Then

        Me.TxtEntrada.Text = "00:00"
        Me.TxtSalida.Text = "00:00"
        Me.TxtEntrada1.Text = "00:00"
        Me.TxtEntrada2.Text = "00:00"
        Me.TxtSalida1.Text = "00:00"
        Me.TxtSalida2.Text = "00:00"
        Me.ChkUnicoHorario.Value = xtpUnchecked
  Else
'        Me.TxtHora.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo")
        Me.TxtEntrada.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo")
        Me.TxtSalida.Text = Me.AdoConsulta.Recordset("SalidaAlmuerzo")
        Me.TxtEntrada1.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo1")
        Me.TxtEntrada2.Text = Me.AdoConsulta.Recordset("EntradaAlmuerzo2")
        Me.TxtSalida1.Text = Me.AdoConsulta.Recordset("SalidaAlmuerzo1")
        Me.TxtSalida2.Text = Me.AdoConsulta.Recordset("SalidaAlmuerzo2")
        If Me.AdoHorarios2.Recordset("ExcluirSabado") = True Then
          Me.ChkUnicoHorario.Value = xtpChecked
        Else
          Me.ChkUnicoHorario.Value = xtpUnchecked
        End If
        
        If Me.AdoHorarios2.Recordset("PersonalSinHorario") = True Then
          Me.ChkAsignarHorario.Value = xtpChecked
        Else
          Me.ChkAsignarHorario.Value = xtpUnchecked
        End If
  End If
  
End Sub

Private Sub Text2_Change()

End Sub

Private Sub Text1_Change()

End Sub
