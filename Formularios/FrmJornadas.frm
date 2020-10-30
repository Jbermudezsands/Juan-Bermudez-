VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Object = "{CFFE0A60-8E3A-11D3-BCC0-00104B9E0792}#1.0#0"; "ssinput1.ocx"
Begin VB.Form FrmJornadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Jornadas Laborales"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   7800
   Begin VB.Frame Frame4 
      Caption         =   "Definicion de Horarios"
      Height          =   1455
      Left            =   120
      TabIndex        =   22
      Top             =   4440
      Visible         =   0   'False
      Width           =   3735
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJornadas.frx":0000
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
      Begin MSMask.MaskEdBox TxtHoraEntrada1 
         Height          =   285
         Left            =   1080
         TabIndex        =   24
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmJornadas.frx":0070
         TabIndex        =   25
         Top             =   360
         Width           =   855
      End
      Begin MSMask.MaskEdBox TxtHoraEntrada2 
         Height          =   285
         Left            =   2880
         TabIndex        =   26
         Top             =   360
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJornadas.frx":00E0
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
      Begin MSMask.MaskEdBox TxtHoraSalida1 
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   2040
         OleObjectBlob   =   "FrmJornadas.frx":014E
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin MSMask.MaskEdBox TxtHoraSalida2 
         Height          =   285
         Left            =   2880
         TabIndex        =   30
         Top             =   840
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
   End
   Begin MSAdodcLib.Adodc AdoJornada 
      Height          =   375
      Left            =   600
      Top             =   6240
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
      Caption         =   "AdoJornada"
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
   Begin VB.Frame Frame3 
      Caption         =   "Frame3"
      Height          =   1455
      Left            =   3960
      TabIndex        =   13
      Top             =   2640
      Width           =   3735
      Begin XtremeSuiteControls.RadioButton OptIntercalar 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Jornada Intercalada entre dias"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.RadioButton OptJornadaDia 
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   3015
         _Version        =   786432
         _ExtentX        =   5318
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Jornada Iniicia y Termina en un dia"
         UseVisualStyle  =   -1  'True
         Value           =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Rango de Busquedas"
      Height          =   1455
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   3735
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJornadas.frx":01BC
         TabIndex        =   20
         Top             =   840
         Width           =   1575
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJornadas.frx":023A
         TabIndex        =   19
         Top             =   480
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TxtEntrada2 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   840
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
         Left            =   1800
         TabIndex        =   11
         Top             =   480
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
         Caption         =   "Ventanas de Busquedas"
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
         Left            =   0
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   7575
      Begin TrueOleDBList80.TDBCombo TDBCodigo 
         Bindings        =   "FrmJornadas.frx":02B8
         Height          =   315
         Left            =   840
         TabIndex        =   21
         Top             =   360
         Width           =   1695
         _ExtentX        =   2990
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
         ListField       =   "CodigoJornada"
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
         _PropDict       =   $"FrmJornadas.frx":02D1
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmJornadas.frx":037B
         TabIndex        =   18
         Top             =   840
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmJornadas.frx":03E5
         TabIndex        =   17
         Top             =   360
         Width           =   735
      End
      Begin VB.TextBox TxtNombre 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Top             =   840
         Width           =   6375
      End
      Begin Threed.SSCommand CmdBusqueda 
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   300
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   661
         _Version        =   196610
         Font3D          =   1
         PictureFrames   =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "FrmJornadas.frx":044F
         Caption         =   "Buscar"
         ButtonStyle     =   4
         PictureAlignment=   9
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   3960
         OleObjectBlob   =   "FrmJornadas.frx":09E9
         TabIndex        =   31
         Top             =   360
         Width           =   1335
      End
      Begin ActiveInput.SSSpin SSPinNumCuotas 
         Height          =   450
         Left            =   5280
         TabIndex        =   32
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   794
         _Version        =   65536
         Text            =   "8"
         MaxValue        =   24
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   7815
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      Begin VB.Image Image2 
         Height          =   1020
         Left            =   240
         Picture         =   "FrmJornadas.frx":0A65
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   7800
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "JORNADAS LABORALES - EMPLEADOS SIN HORARIO"
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
         Height          =   615
         Left            =   2280
         TabIndex        =   1
         Top             =   240
         Width           =   3840
      End
   End
   Begin Threed.SSCommand CmdCerrar 
      Height          =   585
      Left            =   6600
      TabIndex        =   4
      ToolTipText     =   "Cerrar la ventana de nuevo Crédito"
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1032
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmJornadas.frx":4BC6
      Caption         =   "Cerrar"
      ButtonStyle     =   3
      PictureAlignment=   9
      PictureDnFrames =   1
      PictureDn       =   "FrmJornadas.frx":58A0
   End
   Begin Threed.SSCommand CmdGuardar 
      Height          =   585
      Left            =   5400
      TabIndex        =   5
      ToolTipText     =   "Guarda el Crédito"
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1032
      _Version        =   196610
      CaptionStyle    =   1
      MarqueeDirection=   1
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmJornadas.frx":5B02
      Caption         =   "Guardar  "
      ButtonStyle     =   3
      PictureAlignment=   9
      ShapeSize       =   1
   End
   Begin Threed.SSCommand SSCommand1 
      Height          =   585
      Left            =   6600
      TabIndex        =   6
      ToolTipText     =   "Cerrar la ventana de nuevo Crédito"
      Top             =   4200
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   1032
      _Version        =   196610
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmJornadas.frx":63DC
      Caption         =   "Cerrar"
      ButtonStyle     =   3
      PictureAlignment=   9
      PictureDnFrames =   1
      PictureDn       =   "FrmJornadas.frx":70B6
   End
   Begin Threed.SSCommand SSCommand2 
      Height          =   585
      Left            =   5400
      TabIndex        =   7
      ToolTipText     =   "Guarda el Crédito"
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1032
      _Version        =   196610
      CaptionStyle    =   1
      MarqueeDirection=   1
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmJornadas.frx":7318
      Caption         =   "Guardar  "
      ButtonStyle     =   3
      PictureAlignment=   9
      ShapeSize       =   1
   End
   Begin Threed.SSCommand CmdQuitarCredito 
      Height          =   585
      Left            =   4200
      TabIndex        =   8
      ToolTipText     =   "Quita el Crédito de la lista de pagos"
      Top             =   4200
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1032
      _Version        =   196610
      CaptionStyle    =   1
      MarqueeDirection=   1
      PictureFrames   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "FrmJornadas.frx":7BF2
      Caption         =   "Borrar"
      ButtonStyle     =   3
      PictureAlignment=   9
      ShapeSize       =   1
   End
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   480
      Top             =   6720
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
End
Attribute VB_Name = "FrmJornadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBusqueda_Click()
  Quien = "Jornadas"
  FrmConsultaReloj.Show 1
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdGuardar_Click()
 Me.AdoConsulta.RecordSource = "SELECT Jornada.* From Jornada WHERE (((CodigoJornada)='" & Me.TDBCodigo.Text & "'))"
 Me.AdoConsulta.Refresh
 If Me.AdoConsulta.Recordset.EOF Then
 
   Me.AdoConsulta.Recordset.AddNew
    Me.AdoConsulta.Recordset("CodigoJornada") = Me.TDBCodigo.Text
    Me.AdoConsulta.Recordset("NombreJornada") = Me.TxtNombre.Text
    Me.AdoConsulta.Recordset("HorasLaborales") = Me.SSPinNumCuotas.Text
    Me.AdoConsulta.Recordset("RangoHora1") = Me.TxtEntrada1.Text
    Me.AdoConsulta.Recordset("RangoHora2") = Me.TxtEntrada2.Text
   
    If Me.OptIntercalar.Value = True Then
      Me.AdoConsulta.Recordset("JornadaIntercalada") = True
    Else
      Me.AdoConsulta.Recordset("JornadaIntercalada") = False
    End If
   Me.AdoConsulta.Recordset.Update
 
 Else

     Me.AdoConsulta.Recordset("NombreJornada") = Me.TxtNombre.Text
     Me.AdoConsulta.Recordset("HorasLaborales") = Me.SSPinNumCuotas.Text
     Me.AdoConsulta.Recordset("RangoHora1") = Me.TxtEntrada1.Text
     Me.AdoConsulta.Recordset("RangoHora2") = Me.TxtEntrada2.Text
    
     If Me.OptIntercalar.Value = True Then
       Me.AdoConsulta.Recordset("JornadaIntercalada") = True
     Else
       Me.AdoConsulta.Recordset("JornadaIntercalada") = False
     End If
    Me.AdoConsulta.Recordset.Update
    
 End If
 

     Me.TDBCodigo.Text = ""
    Me.TxtNombre.Text = ""
    Me.SSPinNumCuotas.Text = "8"
    Me.TxtEntrada1.Text = "00:00"
    Me.TxtEntrada2.Text = "00:00"
    Me.OptJornadaDia.Value = True
    
    MsgBox "Registro Grabado con Existo!!", vbExclamation
    
     With Me.AdoJornada
      .ConnectionString = ConexionReloj
      .RecordSource = "SELECT Jornada.* FROM Jornada"
      .Refresh
    End With

    
End Sub

Private Sub CmdQuitarCredito_Click()
Dim Respuesta As Double

 
 Respuesta = MsgBox("Esta Seguro de Borrar el Registro?" & " " & Me.TxtNombre.Text, vbYesNo)
 If Respuesta = 6 Then
        Me.AdoConsulta.RecordSource = "SELECT Jornada.* From Jornada WHERE (((CodigoJornada)='" & Me.TDBCodigo.Text & "'))"
        Me.AdoConsulta.Refresh
        If Not Me.AdoConsulta.Recordset.EOF Then
           Me.AdoConsulta.Recordset.Delete
        End If
    Me.TDBCodigo.Text = ""
    Me.TxtNombre.Text = ""
    Me.SSPinNumCuotas.Text = "8"
    Me.TxtEntrada1.Text = "00:00"
    Me.TxtEntrada2.Text = "00:00"
    Me.OptJornadaDia.Value = True
 End If
 
     MsgBox "Registro Borrado!!", vbInformation
 
    With Me.AdoJornada
      .ConnectionString = ConexionReloj
      .RecordSource = "SELECT Jornada.* FROM Jornada"
      .Refresh
    End With
    
    Me.TDBCodigo.RowSource = Me.AdoJornada
    
    
 
End Sub

Private Sub Form_Activate()
    With Me.AdoJornada
      .ConnectionString = ConexionReloj
      .RecordSource = "SELECT Jornada.* FROM Jornada"
      .Refresh
    End With
    

    
End Sub

Private Sub Form_Load()

With Me.AdoJornada
  .ConnectionString = ConexionReloj
  .RecordSource = "SELECT Jornada.* FROM Jornada"
  .Refresh
End With

With Me.AdoConsulta
  .ConnectionString = ConexionReloj
End With
 

 MDIPrimero.Skin1.ApplySkin hWnd
 
 Me.CmdGuardar.BackColor = RGB(222, 227, 247)
 Me.CmdBusqueda.BackColor = RGB(222, 227, 247)
Me.CmdCerrar.BackColor = RGB(222, 227, 247)
Me.CmdQuitarCredito.BackColor = RGB(222, 227, 247)
Me.TxtEntrada1.Text = "00:00"
Me.TxtEntrada2.Text = "00:00"


End Sub

Private Sub TDBCodigo_Change()
 Me.AdoConsulta.RecordSource = "SELECT Jornada.* From Jornada WHERE (((CodigoJornada)='" & Me.TDBCodigo.Text & "'))"
 Me.AdoConsulta.Refresh
 If Not Me.AdoConsulta.Recordset.EOF Then
   Me.TxtNombre.Text = Me.AdoConsulta.Recordset("NombreJornada")
   Me.SSPinNumCuotas.Text = Me.AdoConsulta.Recordset("HorasLaborales")
   Me.TxtEntrada1.Text = Me.AdoConsulta.Recordset("RangoHora1")
   Me.TxtEntrada2.Text = Me.AdoConsulta.Recordset("RangoHora2")
   
   If Me.AdoConsulta.Recordset("JornadaIntercalada") = True Then
     Me.OptIntercalar.Value = True
   Else
     Me.OptJornadaDia.Value = True
   End If
Else
    Me.TxtNombre.Text = ""
    Me.SSPinNumCuotas.Text = "8"
    Me.TxtEntrada1.Text = "00:00"
    Me.TxtEntrada2.Text = "00:00"
    Me.OptJornadaDia.Value = True

 End If
End Sub

Private Sub TDBCodigo_ItemChange()
 Me.AdoConsulta.RecordSource = "SELECT Jornada.* From Jornada WHERE (((CodigoJornada)='" & Me.TDBCodigo.Text & "'))"
 Me.AdoConsulta.Refresh
 If Not Me.AdoConsulta.Recordset.EOF Then
   Me.TxtNombre.Text = Me.AdoConsulta.Recordset("NombreJornada")
   Me.SSPinNumCuotas.Text = Me.AdoConsulta.Recordset("HorasLaborales")
   Me.TxtEntrada1.Text = Me.AdoConsulta.Recordset("RangoHora1")
   Me.TxtEntrada2.Text = Me.AdoConsulta.Recordset("RangoHora2")
   
   If Me.AdoConsulta.Recordset("JornadaIntercalada") = True Then
     Me.OptIntercalar.Value = True
   Else
     Me.OptJornadaDia.Value = True
   End If
Else

    Me.TxtNombre.Text = ""
    Me.SSPinNumCuotas.Text = "8"
    Me.TxtEntrada1.Text = "00:00"
    Me.TxtEntrada2.Text = "00:00"
    Me.OptJornadaDia.Value = True


 End If
End Sub
