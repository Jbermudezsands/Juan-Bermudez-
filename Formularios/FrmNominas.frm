VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmNominasReloj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transferencia de Datos Registro de Nominas"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   7830
   Begin VB.CheckBox ChkTodosDptos 
      Caption         =   "Incluir Todos los Departamentos"
      Height          =   315
      Left            =   4200
      TabIndex        =   17
      Top             =   6960
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CheckBox ChkAcumulado 
      Caption         =   "Calcular Acumulado Rango de Fechas"
      Height          =   315
      Left            =   4200
      TabIndex        =   16
      Top             =   6480
      Visible         =   0   'False
      Width           =   4695
   End
   Begin MSAdodcLib.Adodc AdoConexion 
      Height          =   330
      Left            =   360
      Top             =   4800
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
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
   Begin MSAdodcLib.Adodc AdoTipoNominas 
      Height          =   375
      Left            =   360
      Top             =   4320
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "AdoTipoNominas"
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C1A1&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9975
      TabIndex        =   12
      Top             =   0
      Width           =   9975
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Transferencia de Datos"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   2160
         TabIndex        =   13
         Top             =   360
         Width           =   3840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   9960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image2 
         Height          =   960
         Left            =   360
         Picture         =   "FrmNominas.frx":0000
         Stretch         =   -1  'True
         Top             =   80
         Width           =   1080
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton CmdIniciar 
      Caption         =   "Iniciar"
      Height          =   375
      Left            =   5040
      TabIndex        =   6
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7575
      Begin VB.TextBox TxtAnteponer 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   24
         Text            =   "00000"
         Top             =   1320
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox Check5 
         Caption         =   "Horas Turno"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   840
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check4 
         Caption         =   "Anteponer Ceros"
         Height          =   255
         Left            =   4920
         TabIndex        =   22
         Top             =   1320
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Transferir E/S"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox TxtMinutos 
         Enabled         =   0   'False
         Height          =   375
         Left            =   7080
         TabIndex        =   20
         Text            =   "15"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox TxtNumeroNomina 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3480
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "FrmNominas.frx":C1C2
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin TrueOleDBList80.TDBCombo TDBTipo 
         Bindings        =   "FrmNominas.frx":C23A
         DataSource      =   "AdoTipoNominas"
         Height          =   315
         Left            =   2880
         TabIndex        =   15
         Top             =   200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   10583
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
         ComboStyle      =   0
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
         ListField       =   ""
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
         DropdownPosition=   1
         Locked          =   0   'False
         ScrollTrack     =   0   'False
         RowDividerColor =   14215660
         RowSubDividerColor=   14215660
         AddItemSeparator=   ";"
         _PropDict       =   $"FrmNominas.frx":C257
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
      Begin VB.CheckBox ChkNumeroTarjeta 
         Caption         =   "Utilizar Numero Tarjeta"
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   960
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   4800
         OleObjectBlob   =   "FrmNominas.frx":C301
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Horas Laboradas"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Horas  Extras"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker DtpFechaINI 
         Height          =   300
         Left            =   2640
         TabIndex        =   4
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   529
         _Version        =   393216
         Format          =   78118913
         CurrentDate     =   40789
      End
      Begin MSComCtl2.DTPicker DTFechaFin 
         Height          =   300
         Left            =   5160
         TabIndex        =   5
         Top             =   600
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   529
         _Version        =   393216
         Format          =   78118913
         CurrentDate     =   40789
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "FrmNominas.frx":C365
         TabIndex        =   9
         Top             =   600
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   2160
         OleObjectBlob   =   "FrmNominas.frx":C3CF
         TabIndex        =   14
         Top             =   240
         Width           =   735
      End
   End
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   7575
      _Version        =   786432
      _ExtentX        =   13361
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar osProgress2 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Visible         =   0   'False
      Width           =   4815
      _Version        =   786432
      _ExtentX        =   8493
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   14737632
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   360
      Top             =   5160
      Width           =   2655
      _ExtentX        =   4683
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
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   3840
      Top             =   4320
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
   Begin MSAdodcLib.Adodc AdoReportes 
      Height          =   375
      Left            =   3840
      Top             =   4800
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
      Caption         =   "AdoReportes"
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
   Begin MSAdodcLib.Adodc AdoHorarios 
      Height          =   375
      Left            =   3840
      Top             =   5400
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
   Begin MSAdodcLib.Adodc AdoBuscaReporte 
      Height          =   375
      Left            =   3960
      Top             =   5880
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
      Caption         =   "AdoBuscaReporte"
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
   Begin MSAdodcLib.Adodc AdoConsultaNomina 
      Height          =   375
      Left            =   360
      Top             =   5760
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "AdoConsultaNomina"
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
   Begin MSAdodcLib.Adodc AdoConsultaEasy 
      Height          =   375
      Left            =   360
      Top             =   6240
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "AdoConsultaEasy"
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
   Begin MSAdodcLib.Adodc AdoEmpleadosNomina 
      Height          =   375
      Left            =   360
      Top             =   6840
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
      Caption         =   "AdoEmpleadosNomina"
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
   Begin MSAdodcLib.Adodc AdoDetalleHoraNomina 
      Height          =   375
      Left            =   360
      Top             =   7200
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
      Caption         =   "AdoDetalleHoraNomina"
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
   Begin MSAdodcLib.Adodc AdoHorasExtraNomina 
      Height          =   375
      Left            =   360
      Top             =   7560
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
      Caption         =   "AdoHorasExtraNomina"
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
   Begin MSAdodcLib.Adodc AdoHorasNomina 
      Height          =   375
      Left            =   360
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
      Caption         =   "AdoHorasNomina"
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
   Begin MSAdodcLib.Adodc AdoHorasTurnoNomina 
      Height          =   375
      Left            =   360
      Top             =   8280
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
      Caption         =   "AdoEmpleadosNomina"
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
Attribute VB_Name = "FrmNominasReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ConexionStringNomina As String


Private Sub Check3_Click()
 If Me.Check3.Value = 0 Then
   Me.Check1.Value = 1
   Me.Check2.Value = 1
   Me.Check5.Value = 1
   Me.Check1.Enabled = True
   Me.Check2.Enabled = True
   Me.Check5.Enabled = True
 Else
   Me.Check1.Value = 0
   Me.Check2.Value = 0
   Me.Check5.Value = 0
   Me.Check1.Enabled = False
   Me.Check2.Enabled = False
   Me.Check5.Enabled = False
 End If


End Sub

Private Sub Check4_Click()
If Me.Check4.Value = 0 Then
  Me.ChkNumeroTarjeta.Value = 1
  Me.ChkNumeroTarjeta.Enabled = True
  Me.TxtAnteponer.Visible = False
Else
  Me.ChkNumeroTarjeta.Value = 0
  Me.ChkNumeroTarjeta.Enabled = False
  Me.TxtAnteponer.Visible = True
  
End If
End Sub

Private Sub CmdIniciar_Click()
 Dim Date1 As Date, Date2 As Date, CodEmpleado As String, CodigoEmpleado As Double, TarifaHoraria As Double
 Dim Ciclo As Double, HorasTrabajadas As String, NumeroNomina As Double, NumeroLinea As Double
 Dim Entrada As String, Salida As String, HorasExtras As Double, HoraSalida As Date, HoraSalidaHorario As Date
 Dim Horas As String, fPreview As New FrmPreview, Dia As Double, TotalHoras As Double, MinutosExtra As Double
 Dim cn As New ADODB.Connection, rpt As Object, CardNumero As String, TotalHorasExtras As Double
 Dim rs As New ADODB.Recordset, Id As Double, DiaInicio As Double, Numero As Double, CodigoInterno As Double
 Dim Fecha As String, HoraHorarioEntrada As Date, HoraLab As Double, Salida2 As Date
 Dim Formato As String

      Formato = Me.TxtAnteponer.Text & "#"
      
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
       If Me.TDBTipo.Text = "" Then
         MsgBox "Seleccione un Tipo de NOmina", vbCritical, "Zeus Reloj"
         Exit Sub
       End If

      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM DetalleHorasProduccion WHERE (NumNomina = " & Me.TxtNumeroNomina.Text & ")", ConexionStringNomina
       rs.Open "DELETE FROM HorasExtras Where (NumNomina = " & Me.TxtNumeroNomina.Text & ")", ConexionStringNomina


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************


       sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"  'DESC
       Me.AdoEmpleados.ConnectionString = ConexionEasy
             
'      sql = "SELECT * From Empleado WHERE (CodTipoNomina = '" & Me.TDBTipo.Text & "') AND (Activo = 1)"
'      Me.AdoEmpleados.ConnectionString = ConexionStringNomina

      
      Me.AdoEmpleados.RecordSource = sql
      Me.AdoEmpleados.Refresh
      If Not Me.AdoEmpleados.Recordset.EOF Then
        Me.AdoEmpleados.Recordset.MoveLast
        Me.osProgress1.Max = Me.AdoEmpleados.Recordset.RecordCount
      Else
         Me.osProgress1.Max = 0
      End If
      Me.osProgress1.Min = 0
      Me.osProgress1.Value = 0
      i = 0
      Me.osProgress1.Visible = True
      
      If Not Me.AdoEmpleados.Recordset.BOF Then
       Me.AdoEmpleados.Recordset.MoveFirst
      End If
      Me.AdoReportes.RecordSource = "SELECT Reportes.* FROM Reportes "
      Me.AdoReportes.Refresh
      
     
     

      Do While Not Me.AdoEmpleados.Recordset.EOF
        DoEvents
        
        CodEmpleado = Me.AdoEmpleados.Recordset("Userid")
'        If Me.AdoEmpleados.Recordset("CodEmpleado1") <> "" Then
'          CodEmpleado = CDbl(Me.AdoEmpleados.Recordset("CodEmpleado1"))
'        Else
'          CodEmpleado = 0
'        End If
        Dia = 1
        
        CodigoH = ""
        TieneJornadas = False
        
        Me.osProgress2.Min = 0
        Me.osProgress2.Max = DateDiff("d", Me.DTPFechaIni.Value, Me.DTFechaFin.Value)
        Me.osProgress2.Value = 0
        Me.osProgress2.Visible = True
        
        
        Contador = 0
        FechaInicial = Me.DTPFechaIni.Value
        Do While FechaInicial <= DTFechaFin.Value
         Me.Caption = "Procesando " & FechaInicial & " Empleado: " & i & " de " & Me.osProgress1.Max
         DoEvents
         
         '********************************************************************************************
         '///////////////CON ESTA CONSULTA BUSCO LOS DATOS DE CONFIGURACION //////////////////////////
         '********************************************************************************************
           MDIPrimero.DtaEmpresa.Refresh
           If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
           
             
             DiaExtra = DiaSemana(Day(FechaInicial), Month(FechaInicial), Year(FechaInicial))
             If DiaExtra = 6 Then
              ConfHorasTrabajadas = MDIPrimero.DtaEmpresa.Recordset("HorasTrabSab")
             ElseIf DiaExtra = 0 Then
              ConfHorasTrabajadas = MDIPrimero.DtaEmpresa.Recordset("HorasTrabDom")
             Else
              ConfHorasTrabajadas = MDIPrimero.DtaEmpresa.Recordset("HorasTrab")
             End If
             ConfCalcularHorasTrab = MDIPrimero.DtaEmpresa.Recordset("CalcularHorasTrab")
           End If

                '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                      "WHERE ((UserShift.Userid)='" & CodEmpleado & "')"
                Me.AdoConsulta.RecordSource = sql
                Me.AdoConsulta.Refresh
                If Not Me.AdoConsulta.Recordset.EOF Then
                  FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                  Ciclo = Me.AdoConsulta.Recordset("Cycles")
                  Date1 = CDate(FechaInicioH)
                  Date2 = CDate(FechaInicial)  'Me.DtpFechaINI.Value
                  DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                 
                End If

        
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                '////////////BUSCO EL HORARIO PARA ESTE EMPLEADO ////////////////////////////////////////////////////////////////
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                 Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                    If Not Me.AdoHorarios.Recordset.EOF Then
                    
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SIGNIFICA QUE TIENE HORARIO PERO NO PARA ESTE DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        

                            Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                          "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                            Me.AdoHorarios.Refresh
                            
                           LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                          If LongitudMinutosIn < 1200 Then
                          
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                             
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                          
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                        SinHorario = True
                      Else
                        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '/////////////////////SIGNICA QUE TIENE HORARIO Y TAMBIEN TIENE ASIGIINADO PARA ESTE DIA ///////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////
                        SinHorario = False
                        
                            '********************************************************************************************
                            '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                            '********************************************************************************************
                            If Not Me.AdoHorarios.Recordset.EOF Then
                              CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                              CodigoH = Me.AdoHorarios.Recordset("Schid")
                            End If
                            Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
                            Me.AdoBuscaReporte.Refresh
                            If Not Me.AdoBuscaReporte.Recordset.EOF Then
                            '/////SI TIENE HORAS EXTRA EN EL HORARIO, SE CAMBIA LA CONFIGURACION GENERAL ////////////
                            TipoHorasTrabajada = Me.AdoBuscaReporte.Recordset("TipoCalcularHorasTrab")
                            DiaExtra = DiaSemana(Day(FechaInicial), Month(FechaInicial), Year(FechaInicial))
                            If DiaExtra = 6 Then
                               ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrabSab")
                            ElseIf DiaExtra = 0 Then
                               ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrabDom")
                            Else
                               ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrab")
                            End If
                               ConfCalcularHorasTrab = Me.AdoBuscaReporte.Recordset("CalcularHorasTrab")
    
                            End If
                            
    
                        
                         TieneJornadas = False
                         
                           BInTime = Me.AdoHorarios.Recordset("BIntime")
                           EInTime = Me.AdoHorarios.Recordset("EIntime")
                           InTime = Me.AdoHorarios.Recordset("Intime")
                           LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                           Me.AdoHorarios.Recordset.MoveLast
                           
                           BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                           EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                           OutTime = Me.AdoHorarios.Recordset("OutTime")
                           LongitudMinutosOut = Me.AdoHorarios.Recordset("Longtime")
                           TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                           
                           FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                           MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                           MinutosTarde = MinutosSalida & ":00" & ":00"
                           FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                           FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                           
            
            
                           
                           sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
            
                           
                           '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                           '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                           '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    
                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                           MinutosTarde = MinutosSalida & ":00" & ":00"
                           FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                           FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                           
    '                       HorasIn = Int(LongitudMinutosIn / 60) & ":" & Int(LongitudMinutosIn Mod 60)
    
    '                       If (CDate(InTime) + CDate(HorasIn)) > CDate(Fecha) Then
    '                        '////SI LA SALIDA ES PARA EL DIA SIGUIENTE PASO PARA EL DIA SIGUIENTE
    '                        FechaHInicio = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & BOutTime & "#"
    '                        FechaHFinal = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & EOutTime & "#" '+ CDate(MinutosTarde) 'Me.DTFechaFin.Value
    ''                        FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
    '                       End If
                           
                           SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                                       "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                          
                          End If
                      
                      
                Else '//////SI NO TIENE HORARIO SOLO AGREGO LOS REGISTROS DE ENTRADA ///////////
                
                        FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & "#"
                        FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59#"
                       
                       BInTime = "?"
                       EInTime = "?"
                       InTime = "?"
                       
        '               Me.AdoHorarios.Recordset.MoveLast
                       
                       BOutTime = "?"
                       EOutTime = "?"
                       OutTime = "?"
                       
'                       FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & "#"  'Me.DtpFechaINI.Value
'                       FechaHFinal = CDate(FechaInicial)
'                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " 23:59:59#"
                       

                      
                      '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                      Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                      Me.AdoBuscaReporte.Refresh
                      If Not Me.AdoBuscaReporte.Recordset.EOF Then
                          CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                          HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                          RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                          RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                          JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                          
                         
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                      End If
                    
                    
                    End If
                    
                    

        
        
                    '*********************************************************************************************
                    '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                    '*********************************************************************************************
           
                    Entrada = "00:00"
                    If TieneJornadas = True Then
                    
                        Me.AdoConsulta.RecordSource = sql
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                          Entrada = Me.AdoConsulta.Recordset("CheckTime")
                        End If
                   
                    Else
                        Me.AdoConsulta.RecordSource = sql
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                          Entrada = Me.AdoConsulta.Recordset("CheckTime")
                        End If
                    End If
                    
                    
                   
                    '*********************************************************************************************
                    '///////////////CON ESTA CONSULTA BUSCO LA HORA DE SALIDA///////////////////////////////////
                    '*********************************************************************************************
                    
                        If CodEmpleado = 6 Then
                           CodEmpleado = 6
                        End If
                    
                    Salida = "00:00"
                    If TieneJornadas = True Then
                       
                         '///////////////////////////////CON ESTAS FECHAS BUSCO LA HORA DE SALIDA DE LA JORNADA ///////////////////
                         
                         
                         HoraSalida = CDate(Entrada) + CDate(CInt(HorasLaborales) & ":00:00")
                         FechaHInicio = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") - CDate(RangoHora1 & ":00")), "mm/dd/yyyy hh:mm:ss")
                         FechaHFinal = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") + CDate(RangoHora2 & ":00")), "mm/dd/yyyy hh:mm:ss")
                         HoraSalida = Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59"
                         HoraSalida = Format(HoraSalida, "dd/mm/yyyy hh:mm:ss")
                         If JornadaIntercalada = False Then
                            If CDate(FechaHFinal) > CDate(HoraSalida) Then
                               FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                            End If
                         End If
                   
                        FechaHInicio = "#" & FechaHInicio & "#"
                        FechaHFinal = "#" & FechaHFinal & "#"
                        
                        SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                                    "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"

                   
                        Me.AdoConsulta.RecordSource = SQlSalida
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                            Me.AdoConsulta.Recordset.MoveLast
                            Salida = Me.AdoConsulta.Recordset("CheckTime")
                        ElseIf JornadaIntercalada = True Then
                          '//////////////SI LA JORNADA ES INTERCALADA Y NO TIENE REGISTRO DE SALIDA /////////////////////////
                          '//////////////HAGO CERO LA ENTRADA ///////////////////////////////////////////////////////
                            Entrada = "00:00"
                        End If
                   
                    Else
                        Me.AdoConsulta.RecordSource = SQlSalida
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                          Me.AdoConsulta.Recordset.MoveLast
                          Salida = Me.AdoConsulta.Recordset("CheckTime")
                        End If
                    End If
                    
                    '*********************************************************************************************
                    '///////////////CON ESTA CONSULTA BUSCO EL NOMBRE DEL EMPLEADO///////////////////////////////////
                    '*********************************************************************************************
                    sql = "SELECT Userinfo.*, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh

                    If Not Me.AdoConsulta.Recordset.EOF Then
                      If Not IsNull(Me.AdoConsulta.Recordset("Name")) Then
                        NombreEmpleado = Me.AdoConsulta.Recordset("Name")
                      Else
                        NombreEmpleado = ""
                      End If
                      If Not IsNull(Me.AdoConsulta.Recordset("DeptName")) Then
                       departamento = Me.AdoConsulta.Recordset("DeptName")
                      End If
                              
                      If Me.ChkNumeroTarjeta.Value = 1 Then
                        If Not IsNull(Me.AdoConsulta.Recordset("Cardnum")) Then
                          If Me.AdoConsulta.Recordset("Cardnum") <> "" Then
                            CardNumero = Me.AdoConsulta.Recordset("Cardnum")
                          End If
                        End If
                      Else
                        Me.AdoEmpleadosNomina.RecordSource = "SELECT  * From Empleado WHERE (CodEmpleado1 = " & CodEmpleado & ") AND (Activo = 1)"
                        Me.AdoEmpleadosNomina.Refresh
                        If Not Me.AdoEmpleadosNomina.Recordset.EOF Then
                         '-----------------BUSCO EL CODIGO INTERNO ------------------------------------
                         Fecha = Format(CDate(FechaInicial), "yyyy-mm-dd")
                         CodigoInterno = Me.AdoEmpleadosNomina.Recordset("CodEmpleado")
                         CardNumero = CodEmpleado
                        End If
                      
                      
                      End If
                      
                    End If

                     
                    
                    '*********************************************************************************************
                    '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                    '*********************************************************************************************
                  If InTime <> "?" Then
                    If Entrada <> "00:00" Then
                      If ConfCalcularHorasTrab = True Then
                          If TipoHorasTrabajada = "HorasTrab" Then
                             If InTime > Format(Entrada, "hh:mm") Then
                                Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                             End If
                          End If
                      End If
                    End If
                  End If

                    
                    HorasTrabajadas = 0
                    If Salida <> "00:00" Then
                     If Entrada <> "00:00" Then
'                      HorasTrabajadas = (DateDiff("n", Entrada, Salida)) / 60
'                      HoraLaboradas = ConvertirSegundos((DateDiff("s", Entrada, Salida)), DiaInicio)
                      HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60) - RestarAlmuerzo
                      
                      
'                       HorasTrabajadas = ConvertirSegundos((DateDiff("s", Entrada, Salida)))
                      HoraSalida = Format(Salida, "hh:mm:ss")
                     Else
                      HorasTrabajadas = 0
                     End If
                    End If
                    
                    HorasExtras = 0
                    Horas = "0:00"
                                              
                    HoraHorarioEntrada = "07:00:00"
                     If InTime <> "?" Then
                        If InTime <> "" Then
                            HoraHorarioEntrada = CDate(InTime)
                         Else
                            HoraHorarioEntrada = "07:00:00"
                         End If
                      Else
                        HoraHorarioEntrada = "07:00:00"
                      End If
                       
                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                       

                    
                        If Salida <> "00:00" Then
                        
                                If OutTime <> "?" Then
                                  If OutTime <> "" Then
                                    HoraSalidaHorario = OutTime
                                  Else
                                    HoraSalidaHorario = "17:30:00"
                                  End If
                                Else
                                   HoraSalidaHorario = "17:30:00"
                                End If
                           If Entrada <> "00:00" Then
                           
                           
                            '***********************************************************************************
                            '//////////////VERIFICO SI LAS HORAS EXTRAS SE CALCULAN POR HORAS TRABAJADAS ///////
                            '***********************************************************************************
                            If TieneJornadas = True Then
                               If CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) > HorasLaborales Then
                                   HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                   Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                   HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales)
                               End If
                            Else
                                If ConfCalcularHorasTrab = False Then
                                  If SinHorario = False Then
                                   Horas = ConvertirSegundos((DateDiff("s", HoraSalidaHorario, HoraSalida)), DiaInicio)
                                   HorasExtras = (CDbl(((DateDiff("n", HoraSalidaHorario, HoraSalida)) / 60)))
                                  Else
                                   HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600))) * 3600
                                   Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                   HorasExtras = (CDbl(((DateDiff("n", Entrada, Salida)) / 60)))
                                  End If
                                ElseIf CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) > ConfHorasTrabajadas Then
                                   HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) - ConfHorasTrabajadas) * 3600
                                   Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                   HorasExtras = (CDbl(((DateDiff("n", Entrada, Salida)) / 60) - RestarAlmuerzo) - ConfHorasTrabajadas)
                                End If
                            End If
                            
                            
                         Else
                             HorasExtras = 0
                         End If
                        Else
                         HorasExtras = 0
                       End If
                        
                    If HorasExtras < 0 Then
                      HorasExtras = 0
                    End If
                        
                    '--------------------------------------------------------------------------------------------------------------------------------------------------------
                    '--------------------------------------------RESTO EL TOTAL DE HORAS EXTRAS DE LAS LABORADAS ------------------------------------------------------------
                    '--------------------------------------------------------------------------------------------------------------------------------------------------------
                    
                     
                   
                    
                    If Me.TxtMinutos.Text <> "" Then
                     If IsNumeric(Me.TxtMinutos.Text) Then
                      MinutosExtra = CDbl(Me.TxtMinutos.Text) / 60
                     
                      If MinutosExtra > HorasExtras Then
                         HorasExtras = 0
                      End If
                     
                     End If
                    End If
                    
                    HorasTrabajadas = CDbl(HorasTrabajadas) - CDbl(HorasExtras)
                  If HorasTrabajadas < 0 Then
                    HorasTrabajadas = 0
                  End If
                    
                   '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    '////////////////////////SI NO QUIERE ENTRADAS Y SALIDAS///////////////////////////////////////
                    '/////////////////////////////////////////////////////////////////////////////////////
             If Me.Check3.Value = 0 Then
                    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                    '/////////////////////////////////////////////////////////////////////////////////////
                  

                  
                  
                  If Me.ChkAcumulado.Value = 0 Then
                  
                                    If Me.ChkNumeroTarjeta.Value = 1 Then
                                      Me.AdoEmpleadosNomina.RecordSource = "SELECT  * From Empleado WHERE (CodEmpleado1 = '" & CardNumero & "') AND (Activo = 1)"
                                    Else
                                      If Me.Check4.Value = 1 Then
                                        Numero = CodEmpleado
                                        CardNumero = Format(Numero, "000#")
                                      End If
                                      Me.AdoEmpleadosNomina.RecordSource = "SELECT  * From Empleado WHERE (CodEmpleado1 = '" & CardNumero & "') AND (Activo = 1)"
                                    End If
                          
                         Me.AdoEmpleadosNomina.RecordSource = "SELECT  * From Empleado WHERE (CodEmpleado1 = '" & CardNumero & "')"
                         Me.AdoEmpleadosNomina.Refresh
                         
                         If Not Me.AdoEmpleadosNomina.Recordset.EOF Then
                         
                           CodigoEmpleado = Me.AdoEmpleadosNomina.Recordset("CodEmpleado")
                           NumeroNomina = Me.TxtNumeroNomina.Text
                           
                           '------------------------BORROR EL REGISTRO DE HORAS EN LA  PLANILLA -----------------------
'                           rs.Open "DELETE FROM DetalleHorasProduccion Where (CodEmpleado = " & CodigoEmpleado & ") And (NumNomina = " & NumeroNomina & ")", ConexionStringNomina
'                           rs.Open "DELETE FROM HorasExtras Where (CodEmpleado = " & CodigoEmpleado & ") And (NumNomina = " & NumeroNomina & ")", ConexionStringNomina
                           
                           If Not IsNull(Me.AdoEmpleadosNomina.Recordset("TarifaHoraria")) Then
                             TarifaHoraria = Me.AdoEmpleadosNomina.Recordset("TarifaHoraria")
                           Else
                             TarifaHoraria = 0
                           End If
                           Me.AdoConsultaNomina.RecordSource = "SELECT  * From DetalleHorasProduccion ORDER BY NumLinea"
                           Me.AdoConsultaNomina.Refresh
                           If Me.AdoConsultaNomina.Recordset.EOF Then
                             NumeroLinea = 1
                           Else
                             Me.AdoConsultaNomina.Recordset.MoveLast
                             NumeroLinea = Me.AdoConsultaNomina.Recordset("NumLinea") + 1
                           End If
                           
                           '/////////////////////BUSCO DETALLE HORA TRABAJADA //////////////////////////////
                           Me.AdoDetalleHoraNomina.RecordSource = "SELECT  * From DetalleHorasProduccion Where (CodEmpleado = " & CodigoEmpleado & ") And (NumNomina = " & NumeroNomina & ")"
                           Me.AdoDetalleHoraNomina.Refresh
                           If Me.AdoDetalleHoraNomina.Recordset.EOF Then
                              Me.AdoDetalleHoraNomina.Recordset.AddNew
                                TotalHoras = 0
                                Me.AdoDetalleHoraNomina.Recordset("CodEmpleado") = CodigoEmpleado
                                Me.AdoDetalleHoraNomina.Recordset("NumNomina") = NumeroNomina
                                Me.AdoDetalleHoraNomina.Recordset("NumLinea") = NumeroLinea
                                Select Case Dia
                                  Case 1: Me.AdoDetalleHoraNomina.Recordset("Lunes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 2: Me.AdoDetalleHoraNomina.Recordset("Martes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 3: Me.AdoDetalleHoraNomina.Recordset("Miercoles") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 4: Me.AdoDetalleHoraNomina.Recordset("Jueves") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 5: Me.AdoDetalleHoraNomina.Recordset("Viernes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 6: Me.AdoDetalleHoraNomina.Recordset("Sabado") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 7: Me.AdoDetalleHoraNomina.Recordset("Domingo") = Format(HorasTrabajadas, "##,##0.00")
                                End Select
                                TotalHoras = TotalHoras + Format(HorasTrabajadas, "##,##0.00")
                                Me.AdoDetalleHoraNomina.Recordset("TotalHoras") = TotalHoras
                                Me.AdoDetalleHoraNomina.Recordset("SalarioHora") = TarifaHoraria
                                Me.AdoDetalleHoraNomina.Recordset("TotalSalarioHora") = TotalHoras * TarifaHoraria
                              Me.AdoDetalleHoraNomina.Recordset.Update
                           ElseIf Dia < 8 Then
                                Me.AdoDetalleHoraNomina.Recordset.MoveLast
                                Me.AdoDetalleHoraNomina.Recordset("CodEmpleado") = CodigoEmpleado
                                Me.AdoDetalleHoraNomina.Recordset("NumNomina") = NumeroNomina
'                                Me.AdoDetalleHoraNomina.Recordset("NumLinea") = NumeroLinea
                                Select Case Dia
                                  Case 1: Me.AdoDetalleHoraNomina.Recordset("Lunes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 2: Me.AdoDetalleHoraNomina.Recordset("Martes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 3: Me.AdoDetalleHoraNomina.Recordset("Miercoles") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 4: Me.AdoDetalleHoraNomina.Recordset("Jueves") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 5: Me.AdoDetalleHoraNomina.Recordset("Viernes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 6: Me.AdoDetalleHoraNomina.Recordset("Sabado") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 7: Me.AdoDetalleHoraNomina.Recordset("Domingo") = Format(HorasTrabajadas, "##,##0.00")
                                End Select
                                TotalHoras = TotalHoras + Format(HorasTrabajadas, "##,##0.00")
                                Me.AdoDetalleHoraNomina.Recordset("TotalHoras") = TotalHoras
                                Me.AdoDetalleHoraNomina.Recordset("SalarioHora") = TarifaHoraria
                                Me.AdoDetalleHoraNomina.Recordset("TotalSalarioHora") = TotalHoras * TarifaHoraria
                              Me.AdoDetalleHoraNomina.Recordset.Update
                           Else
                                Dia = 1
                                TotalHoras = 0
                                Me.AdoDetalleHoraNomina.Recordset.AddNew
                                Me.AdoDetalleHoraNomina.Recordset("CodEmpleado") = CodigoEmpleado
                                Me.AdoDetalleHoraNomina.Recordset("NumNomina") = NumeroNomina
                                Me.AdoDetalleHoraNomina.Recordset("NumLinea") = NumeroLinea
                                Select Case Dia
                                  Case 1: Me.AdoDetalleHoraNomina.Recordset("Lunes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 2: Me.AdoDetalleHoraNomina.Recordset("Martes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 3: Me.AdoDetalleHoraNomina.Recordset("Miercoles") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 4: Me.AdoDetalleHoraNomina.Recordset("Jueves") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 5: Me.AdoDetalleHoraNomina.Recordset("Viernes") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 6: Me.AdoDetalleHoraNomina.Recordset("Sabado") = Format(HorasTrabajadas, "##,##0.00")
                                  Case 7: Me.AdoDetalleHoraNomina.Recordset("Domingo") = Format(HorasTrabajadas, "##,##0.00")
                                End Select
                                TotalHoras = TotalHoras + Format(HorasTrabajadas, "##,##0.00")
                                Me.AdoDetalleHoraNomina.Recordset("TotalHoras") = TotalHoras
                                Me.AdoDetalleHoraNomina.Recordset("SalarioHora") = TarifaHoraria
                                Me.AdoDetalleHoraNomina.Recordset("TotalSalarioHora") = TotalHoras * TarifaHoraria
                              Me.AdoDetalleHoraNomina.Recordset.Update
                              
                           End If
                        
                           
                           
                               '////////////////////////BUSCO EL CONSECUTIVO DE HORAS EXTRA ////////////////////////////
                               Me.AdoConsultaNomina.RecordSource = "SELECT  * From HorasExtras"
                               Me.AdoConsultaNomina.Refresh
                               If Not Me.AdoConsultaNomina.Recordset.EOF Then
                                 Me.AdoConsultaNomina.Recordset.MoveLast
                                 Id = Me.AdoConsultaNomina.Recordset("Id") + 1
                               Else
                                 Id = 1
                               End If
                               
                               If HorasExtras > 0 Then
                                    Me.AdoHorasExtraNomina.RecordSource = "SELECT  * From HorasExtras Where (CodEmpleado = " & CodigoEmpleado & ") And (NumNomina = " & NumeroNomina & ")"
                                    Me.AdoHorasExtraNomina.Refresh
                                    If Me.AdoHorasExtraNomina.Recordset.EOF Then
                                      Me.AdoHorasExtraNomina.Recordset.AddNew
                                        Me.AdoHorasExtraNomina.Recordset("Id") = Id
                                        Me.AdoHorasExtraNomina.Recordset("CodEmpleado") = CodigoEmpleado
                                        Me.AdoHorasExtraNomina.Recordset("NumNomina") = NumeroNomina
                                        Me.AdoHorasExtraNomina.Recordset("CantHoras") = Format(HorasExtras, "##,##0.00")
                                      Me.AdoHorasExtraNomina.Recordset.Update
                                    Else
                                      Me.AdoHorasExtraNomina.Recordset("CantHoras") = Format(Me.AdoHorasExtraNomina.Recordset("CantHoras") + HorasExtras, "##,##0.00")
                                      Me.AdoHorasExtraNomina.Recordset.Update
                                    End If
                                End If
                                
                                '---------------------------------------------------------------------------------------------------------
                                '------------------------BUSCO EL MONTO POR TURNO -----------------------------------
                                '------------------------------------------------------------------------------------
                                
        If CodEmpleado = "2171" Then
          Cod = 1
        End If
                             
                                If Me.Check5.Value = 1 Then
                                 Me.AdoHorasTurnoNomina.RecordSource = "SELECT * From HorasTurno Where (CodEmpleado = " & CodigoEmpleado & ") And (NumNomina = " & NumeroNomina & ")"
                                 Me.AdoHorasTurnoNomina.Refresh
                                   If Me.AdoHorasTurnoNomina.Recordset.EOF Then
                                      Me.AdoHorasTurnoNomina.Recordset.AddNew
                                      Me.AdoHorasTurnoNomina.Recordset("CodEmpleado") = CodigoEmpleado
                                      Me.AdoHorasTurnoNomina.Recordset("NumNomina") = NumeroNomina
                                      Me.AdoHorasTurnoNomina.Recordset("CantHoras") = TotalHoras
                                      Me.AdoHorasTurnoNomina.Recordset("NTurnos") = TotalHoras / 10
                                      Me.AdoHorasTurnoNomina.Recordset.Update
                                   Else
                                      Me.AdoHorasTurnoNomina.Recordset("CantHoras") = TotalHoras
                                      Me.AdoHorasTurnoNomina.Recordset("NTurnos") = TotalHoras / 10
                                      Me.AdoHorasTurnoNomina.Recordset.Update
                                   End If
                                End If
                             
                        
                          Dia = Dia + 1


                      End If
                 End If
             Else
               '////////////////////////////////////BUSCO SI EL EMPLEADO EXISTEN EN LA NOMINA //////////////////
                  If Me.ChkNumeroTarjeta.Value = 1 Then
                    Me.AdoEmpleadosNomina.RecordSource = "SELECT  * From Empleado WHERE (CodEmpleado1 = '" & CardNumero & "') AND (Activo = 1)"
                  Else
                    If Me.Check4.Value = 1 Then
                      Numero = CodEmpleado
'                      CardNumero = Format(Numero, "00000#")
                      CardNumero = Format(Numero, Formato)
                    End If
                    Me.AdoEmpleadosNomina.RecordSource = "SELECT  * From Empleado WHERE (CodEmpleado1 = '" & CardNumero & "') AND (Activo = 1)"
                  End If
                  
                  If CodEmpleado = "2003" Then
                    CodEmpleado = "2003"
                  End If
                  
                  Me.AdoEmpleadosNomina.Refresh
                  If Not Me.AdoEmpleadosNomina.Recordset.EOF Then
                   '-----------------BUSCO EL CODIGO INTERNO ------------------------------------
                   Fecha = Format(CDate(FechaInicial), "yyyy-mm-dd")
                   CodigoInterno = Me.AdoEmpleadosNomina.Recordset("CodEmpleado")
'                   Me.AdoHorasNomina.RecordSource = "SELECT CodEmpleado1, CodEmpleado, CodTipoNomina, FechaEntrada, HoraEntrada, HoraSalida, FechaSalida, bActivo, CodTurno, HREntrada, HRSalida " & _
'                                                    "FROM AsistenciaEmpleado WHERE CodEmpleado1 ='" & CodEmpleado & "' AND bActivo=1 ORDER BY FechaEntrada DESC"
                   Me.AdoHorasNomina.RecordSource = "SELECT * From AsistenciaEmpleado WHERE (CodEmpleado1 = '" & CardNumero & "')  AND (FechaEntrada = CONVERT(DATETIME, '" & Fecha & "', 102))"
                   Me.AdoHorasNomina.Refresh
                If Me.AdoHorasNomina.Recordset.EOF Then
                
                  If Entrada <> "00:00" Or Salida <> "00:00" Then
                        Me.AdoHorasNomina.Recordset.AddNew
                        Me.AdoHorasNomina.Recordset.Fields("CodEmpleado") = CodigoInterno
                        Me.AdoHorasNomina.Recordset.Fields("CodEmpleado1") = CardNumero
                        Me.AdoHorasNomina.Recordset.Fields("FechaEntrada") = Format(FechaInicial, "dd/mm/yyyy")
                        If Entrada <> "00:00" Then
                         '////////////SI LA HORA DE MARCA ES MARYOR QUE EL HORARIO
                            If HoraHorarioEntrada > Entrada Then
                             Me.AdoHorasNomina.Recordset.Fields("HoraEntrada") = Format(HoraHorarioEntrada, "hh:mm:ss")
                            Else
                             Me.AdoHorasNomina.Recordset.Fields("HoraEntrada") = Format(Entrada, "hh:mm:ss")
                            End If
                         
                         
                          Me.AdoHorasNomina.Recordset.Fields("HREntrada") = Format(Entrada, "hh:mm:ss")
                        Else
                          Me.AdoHorasNomina.Recordset.Fields("HoraEntrada") = Format(Salida, "hh:mm:ss")
                        End If
                        
                        If Salida <> "00:00" Then
                          Me.AdoHorasNomina.Recordset.Fields("HRSalida") = Format(Salida, "hh:mm:ss")
                          Me.AdoHorasNomina.Recordset.Fields("FechaSalida") = Format(FechaInicial, "dd/mm/yyyy")
                          
                          If HoraSalidaHorario > Salida Then
                            Me.AdoHorasNomina.Recordset.Fields("HoraSalida") = Format(HoraSalidaHorario, "hh:mm:ss")
                          Else
                            Me.AdoHorasNomina.Recordset.Fields("HoraSalida") = Format(Salida, "hh:mm:ss")
                          End If
                        End If
                      
                        
                        Me.AdoHorasNomina.Recordset.Fields("CodTurno") = "Diurno"
                        Me.AdoHorasNomina.Recordset.Fields("CodTipoNomina") = Me.AdoEmpleadosNomina.Recordset("CodTipoNomina")
                        Me.AdoHorasNomina.Recordset.Fields("bActivo") = 0
                        Me.AdoHorasNomina.Recordset.Update
                   End If
                 Else

                   If Entrada <> "00:00" Or Salida <> "00:00" Then
                        If Entrada <> "00:00" Then
                          Me.AdoHorasNomina.Recordset.Fields("FechaEntrada") = Format(FechaInicial, "dd/mm/yyyy")
                             '////////////SI LA HORA DE MARCA ES MARYOR QUE EL HORARIO
                            If TimeValue(Format(HoraHorarioEntrada, "hh:mm:ss")) >= TimeValue(Format(Entrada, "hh:mm:ss")) Then
                             Me.AdoHorasNomina.Recordset.Fields("HoraEntrada") = Format(HoraHorarioEntrada, "hh:mm:ss")
                            Else
                             Me.AdoHorasNomina.Recordset.Fields("HoraEntrada") = Format(Entrada, "hh:mm:ss")
                            End If
                            
                          Me.AdoHorasNomina.Recordset.Fields("HREntrada") = Format(Entrada, "hh:mm:ss")
                        Else
                          Me.AdoHorasNomina.Recordset.Fields("HoraEntrada") = Format(Salida, "hh:mm:ss")
                        End If
                        
    
                        If Salida <> "00:00" Then
                         Me.AdoHorasNomina.Recordset.Fields("FechaSalida") = Format(FechaInicial, "dd/mm/yyyy")
                          If TimeValue(Format(HoraSalidaHorario, "hh:mm:ss")) >= TimeValue(Format(Salida, "hh:mm:ss")) Then
                            Me.AdoHorasNomina.Recordset.Fields("HoraSalida") = Format(Salida, "hh:mm")
                          Else
                            HoraLab = DateDiff("n", TimeValue(Format(HoraSalidaHorario, "hh:mm:ss")), TimeValue(Format(Salida, "hh:mm:ss"))) / 60
                            If HoraLab >= 1 Then
                              Salida2 = DateAdd("h", 2, HoraSalidaHorario)
                              Me.AdoHorasNomina.Recordset.Fields("HoraSalida") = Format(Salida2, "hh:mm:ss")
                            Else
                              Me.AdoHorasNomina.Recordset.Fields("HoraSalida") = Format(HoraSalidaHorario, "hh:mm:ss")
                            End If
                          End If
                         Me.AdoHorasNomina.Recordset.Fields("HRSalida") = Format(Salida, "hh:mm:ss")
                        End If
                        Me.AdoHorasNomina.Recordset.Fields("CodTipoNomina") = Me.AdoEmpleadosNomina.Recordset("CodTipoNomina")
                        Me.AdoHorasNomina.Recordset.Fields("bActivo") = 0
                        Me.AdoHorasNomina.Recordset.Update
                   End If
                   
                End If
                   
                  
                  
                  
                  
                  
                  End If
             
             
             
             
             
             
             End If
             
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & FechaInicial & " Empleado: " & i & " de " & Me.osProgress1.Max & " " & CodigoEmpleado
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      

      

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
 Dim Numero As Double, i As Double

 MDIPrimero.Skin1.ApplySkin hWnd
 
 Me.DTPFechaIni.Value = Now
 Me.DTFechaFin.Value = Now
 
 With Me.AdoConexion
   .ConnectionString = ConexionReloj
   .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
   .Refresh
 End With
 
 If Not Me.AdoConexion.Recordset.EOF Then
  Me.TxtMinutos.Text = Me.AdoConexion.Recordset("MinutosExtra")
 End If
 
 ConexionStringNomina = Me.AdoConexion.Recordset("Cadena")
 
  With Me.AdoTipoNominas
   .ConnectionString = ConexionStringNomina
   .RecordSource = "SELECT  CodTipoNomina, Nomina, Periodo From TipoNomina Where (Activa = 1)"
   .Refresh
 End With
 
With Me.AdoConsultaEasy
  .ConnectionString = ConexionEasy
End With


With Me.AdoHorasTurnoNomina
  .ConnectionString = ConexionStringNomina
End With

With Me.AdoHorasExtraNomina
  .ConnectionString = ConexionStringNomina
End With
 
With Me.AdoConsultaNomina
  .ConnectionString = ConexionStringNomina
End With

With Me.AdoEmpleadosNomina
  .ConnectionString = ConexionStringNomina
End With

With Me.AdoDetalleHoraNomina
  .ConnectionString = ConexionStringNomina
End With

With Me.AdoHorasNomina
  .ConnectionString = ConexionStringNomina
End With

With Me.AdoConsulta
  .ConnectionString = ConexionEasy
End With
 
 With Me.AdoEmpleados
  .ConnectionString = ConexionEasy
End With

With Me.AdoHorarios
  .ConnectionString = ConexionEasy
End With

With Me.AdoReportes
  .ConnectionString = ConexionReloj
End With

With Me.AdoBuscaReporte
  .ConnectionString = ConexionReloj
End With

           MDIPrimero.DtaEmpresa.Refresh
           If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
             Numero = MDIPrimero.DtaEmpresa.Recordset("AnteponerCeros")
           Else
             Numero = 5
           End If
           
           Me.TxtAnteponer.Text = ""
           For i = 1 To Numero
            Me.TxtAnteponer.Text = Me.TxtAnteponer.Text & 0
           
           Next

End Sub


Private Sub TDBTipo_ItemChange()
  Me.AdoConsultaNomina.RecordSource = "SELECT  * From Nomina WHERE (Activa = 1) AND (CodTipoNomina = '" & Me.TDBTipo.Text & "')"
  Me.AdoConsultaNomina.Refresh
  If Not Me.AdoConsultaNomina.Recordset.EOF Then
    Me.DTPFechaIni.Value = Me.AdoConsultaNomina.Recordset("FechaNominaINI")
    Me.DTFechaFin.Value = Me.AdoConsultaNomina.Recordset("FechaNomina")
    Me.TxtNumeroNomina.Text = Me.AdoConsultaNomina.Recordset("NumNomina")
  End If
End Sub
