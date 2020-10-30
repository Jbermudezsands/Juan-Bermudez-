VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmExportarReloj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportacion de Archivos"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   14295
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de Registros"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   14055
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExportar.frx":0000
         TabIndex        =   3
         Top             =   300
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPFechaIni 
         Height          =   330
         Left            =   1320
         TabIndex        =   4
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   77987841
         CurrentDate     =   40457
      End
      Begin MSComCtl2.DTPicker DTPFechaFin 
         Height          =   330
         Left            =   5160
         TabIndex        =   5
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   582
         _Version        =   393216
         Format          =   77987841
         CurrentDate     =   40457
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   4080
         OleObjectBlob   =   "FrmExportar.frx":0076
         TabIndex        =   6
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00F5C1A1&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   14295
      TabIndex        =   0
      Top             =   0
      Width           =   14295
      Begin VB.Image Image1 
         Height          =   1020
         Left            =   240
         Picture         =   "FrmExportar.frx":00E6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Exportacion de Archivos"
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
         Left            =   5160
         TabIndex        =   1
         Top             =   360
         Width           =   3345
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   13320
         Y1              =   1080
         Y2              =   1080
      End
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmExportar.frx":1E84
      Height          =   4335
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   7646
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
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
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Movimientos de Indices"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131588"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureCurrentRow(0)=   "bHQAAO4BAABCTe4BAAAAAAAANgAAACgAAAAOAAAACgAAAAEAGAAAAAAAuAEAAAAAAAAAAAAAAAAA"
      PictureCurrentRow(1)=   "AAAAAADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAMbHxgAAAP//"
      PictureCurrentRow(2)=   "/////////////////////////////////////////8bHxgAAxsfGAAAAhIaExsfGxsfGxsfGxsfG"
      PictureCurrentRow(3)=   "xsfGxsfGxsfGxsfGxsfG////xsfGAADGx8YAAACEhoTGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bG"
      PictureCurrentRow(4)=   "x8b////Gx8YAAMbHxgAAAISGhMbHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxsbHxv///8bHxgAAxsfG"
      PictureCurrentRow(5)=   "AAAAhIaExsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfGxsfG////xsfGAADGx8YAAACEhoTGx8bGx8bG"
      PictureCurrentRow(6)=   "x8bGx8bGx8bGx8bGx8bGx8bGx8b////Gx8YAAMbHxgAAAISGhISGhISGhISGhISGhISGhISGhISG"
      PictureCurrentRow(7)=   "hISGhISGhP///8bHxgAAxsfGAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAxsfG"
      PictureCurrentRow(8)=   "AADGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8bGx8YAAA=="
      PictureCurrentRow.vt=   9
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   14215660
      RowDividerColor =   14215660
      RowSubDividerColor=   14215660
      DirectionAfterEnter=   1
      DirectionAfterTab=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
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
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HFFAEFF&,.fgcolor=&H800080&"
      _StyleDefs(20)  =   ":id=22,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(21)  =   ":id=22,.fontname=Lucida Calligraphy"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&HECB877&"
      _StyleDefs(23)  =   ":id=14,.fgcolor=&H800000&,.bold=-1,.fontsize=825,.italic=0,.underline=0"
      _StyleDefs(24)  =   ":id=14,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=14,.fontname=MS Sans Serif"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3,.alignment=2,.bgcolor=&HFF0000&"
      _StyleDefs(27)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(28)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(29)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(30)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(31)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(32)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(33)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(34)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(35)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(36)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(37)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(38)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(43)  =   "Named:id=33:Normal"
      _StyleDefs(44)  =   ":id=33,.parent=0"
      _StyleDefs(45)  =   "Named:id=34:Heading"
      _StyleDefs(46)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   ":id=34,.wraptext=-1"
      _StyleDefs(48)  =   "Named:id=35:Footing"
      _StyleDefs(49)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(50)  =   "Named:id=36:Selected"
      _StyleDefs(51)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(52)  =   "Named:id=37:Caption"
      _StyleDefs(53)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(54)  =   "Named:id=38:HighlightRow"
      _StyleDefs(55)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(56)  =   "Named:id=39:EvenRow"
      _StyleDefs(57)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(58)  =   "Named:id=40:OddRow"
      _StyleDefs(59)  =   ":id=40,.parent=33"
      _StyleDefs(60)  =   "Named:id=41:RecordSelector"
      _StyleDefs(61)  =   ":id=41,.parent=34"
      _StyleDefs(62)  =   "Named:id=42:FilterBar"
      _StyleDefs(63)  =   ":id=42,.parent=33"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DefaultExt      =   "txt"
      FileName        =   "*.txt"
      Filter          =   "txt"
   End
   Begin MSAdodcLib.Adodc AdoDepartamento 
      Height          =   375
      Left            =   6120
      Top             =   8040
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
      Caption         =   "AdoDepartamento"
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
      Left            =   6000
      Top             =   8400
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
      Left            =   1320
      Top             =   8880
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
   Begin MSAdodcLib.Adodc AdoConsulta 
      Height          =   375
      Left            =   1440
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
   Begin MSAdodcLib.Adodc AdoHorarios 
      Height          =   375
      Left            =   1320
      Top             =   9480
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
   Begin MSAdodcLib.Adodc AdoEmpleados2 
      Height          =   375
      Left            =   6120
      Top             =   8760
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
      Caption         =   "AdoEmpleados2"
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
      Left            =   6120
      Top             =   9600
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
   Begin XtremeSuiteControls.ProgressBar osProgress2 
      Height          =   255
      Left            =   4080
      TabIndex        =   8
      Top             =   6960
      Visible         =   0   'False
      Width           =   8655
      _Version        =   786432
      _ExtentX        =   15266
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   14737632
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   1440
      TabIndex        =   9
      Top             =   6480
      Visible         =   0   'False
      Width           =   11295
      _Version        =   786432
      _ExtentX        =   19923
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin SmartButtonProject.SmartButton CmdSalir 
      Height          =   855
      Left            =   12960
      TabIndex        =   10
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Salir"
      Picture         =   "FrmExportar.frx":1E9D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin SmartButtonProject.SmartButton SmartButton1 
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   6360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Exportar"
      Picture         =   "FrmExportar.frx":31AF
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureLayout   =   7
   End
   Begin MSAdodcLib.Adodc AdoExporta 
      Height          =   375
      Left            =   1320
      Top             =   8040
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
      Caption         =   "AdoExporta"
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
Attribute VB_Name = "FrmExportarReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub DtpFechaINI_Change()
  Me.DTPFechaFin.Value = Me.DTPFechaIni.Value
End Sub

Private Sub Form_Load()

Me.DTPFechaFin.Value = Format(Now, "dd/mm/yyyy")
Me.DTPFechaIni.Value = Format(Now, "dd/mm/yyyy")

 Me.DBGTransacciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DBGTransacciones.OddRowStyle.BackColor = &H80000005
 Me.DBGTransacciones.AlternatingRowStyle = True
 
With Me.AdoDepartamento
  .ConnectionString = ConexionEasy
End With

With Me.AdoEmpleados
  .ConnectionString = ConexionEasy
End With

With Me.AdoHorarios
  .ConnectionString = ConexionEasy
End With

With Me.AdoConsulta
  .ConnectionString = ConexionEasy
End With

With Me.AdoReportes
  .ConnectionString = ConexionReloj
End With

With Me.AdoBuscaReporte
  .ConnectionString = ConexionReloj
End With

With Me.AdoEmpleados2
  .ConnectionString = ConexionEasy
End With

With Me.AdoExporta
  .ConnectionString = ConexionReloj
End With





Me.AdoEmpleados2.RecordSource = "SELECT Userinfo.* FROM Userinfo"
Me.AdoEmpleados2.Refresh


Me.AdoDepartamento.RecordSource = "SELECT Dept.Deptid, Dept.DeptName FROM Dept"
Me.AdoDepartamento.Refresh

End Sub

Private Sub PushButton1_Click()

     
      
      
      
End Sub

Private Sub SmartButton1_Click()
Dim sql As String, CodDptoIni As String, CodDptoFin As String
Dim rpt As Object, FechaIni As String, FechaFin As String, CodEmpleado As String, NombreEmpleado As String, departamento As String
Dim fPreview As New FrmPreview, i As Double, Dia As String, FechaInicioH As String, Date1 As Date, Date2 As Date
Dim cn As New ADODB.Connection, DiferenciaDias As Double, DiasCiclo As Double, Periodo As Double, DiaPeriodo As Double
Dim rs As New ADODB.Recordset, FechaActual As Date, DiasSumar As Double, FechaHorario As Date
Dim DiaInicio As Double, Ciclo As Double, BInTime As String, EInTime As String, BOutTime As String, EOutTime As String, TardePermintido As Double, InTime As String, OutTime As String
Dim Entrada As String, Salida As String, HorasTrabajadas As String, HorasExtras As Double, HoraSalida As Date, HoraSalidaHorario As Date
Dim HoraEntrada As Date, HoraHorario As Date, MinutosTarde As String, Cod As Double, FechaIn As String, FechaOut As String
Dim FechaHInicio As String, FechaHFinal As String, SQlSalida As String, j As Double, b As Double, HoraLaboradas As Date
Dim TotalHorasTrabajadas As Double, TotalHorasExtras As Double, HorasTarde As Double, TotalHoras As Double, HoraHorarioSalida As Date, HoraAnticipada As Double
Dim MinutosSalida As Double, LongitudMinutosIn As Double, LongitudMinutosOut As Double
Dim FechaInicial As Date, Contador As Double, HorasMinutos As Date, ConfHorasTrabajadas As Double, ConfCalcularHorasTrab As Boolean
Dim SQLExporta As String, Longitud As Integer, Respuesta As Integer
Dim Cadena As String
Dim TextoMonto As String, TipoMovimiento As String, Maximo As Double, Registros As Double

      FechaIni = "#" & Format(DTPFechaIni, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(DTPFechaFin, "mm/dd/yyyy") & " 23:59:59#"
      
     '********************************************************************************************
     '///////////////CON ESTA CONSULTA BUSCO LOS DATOS DE CONFIGURACION //////////////////////////
     '********************************************************************************************
           MDIPrimero.DtaEmpresa.Refresh
           If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
             ConfHorasTrabajadas = MDIPrimero.DtaEmpresa.Recordset("HorasTrab")
             ConfCalcularHorasTrab = MDIPrimero.DtaEmpresa.Recordset("CalcularHorasTrab")
           End If
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      
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
        Contador = 0
        FechaInicial = Me.DTPFechaIni.Value
        Do While FechaInicial <= Me.DTPFechaFin.Value
         DoEvents

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
                
                '///////////CALCULO EL NUMERO DE DIAS ENTRE HORARIO Y SELECCIONADA ///////////////
                ' Diferencias en dias
                'DateDiff("d", "01/01/2000 14:39:00","01/01/2006 14:00:00")
                'Diferencias en horas
                'DateDiff("h", "01/01/2000 14:39:00","01/01/2006 14:00:00")
                'Diferencias en minutos
                'DateDiff("n", "01/01/2000 14:39:00","01/01/2006 14:00:00")
        '        Date1 = Format(CDate(FechaInicioH), "dd/mm/yyyy")
        '        Date2 = Format(CDate(Me.DtpFechaINI.Value), "dd/mm/yyyy")
        
                
                If CodEmpleado = 48 Then
                  Cod = 1
                End If
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                '////////////BUSCO EL HORARIO PARA ESTE EMPLEADO ////////////////////////////////////////////////////////////////
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                              "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "')) "
        
                Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                    If Not Me.AdoHorarios.Recordset.EOF Then
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
        
                       
                       
                       FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                       MinutosTarde = MinutosSalida & ":00" & ":00"
                       FechaHFinal = CDate(FechaInicial & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                       
                       
                       SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                             
                             
        
                    Else '//////SI NO TIENE HORARIO SOLO AGREGO LOS REGISTROS DE ENTRADA ///////////
                
                        FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & "#"
                        FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59:59#"
                       
                       BInTime = "?"
                       EInTime = "?"
                       InTime = "?"
                       
        '               Me.AdoHorarios.Recordset.MoveLast
                       
                       BOutTime = "?"
                       EOutTime = "?"
                       OutTime = "?"
                      sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                      "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                    
                      SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                      "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                    End If
                    
                    

        
        
                    '*********************************************************************************************
                    '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                    '*********************************************************************************************
            
                    Entrada = "00:00"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      Entrada = Me.AdoConsulta.Recordset("CheckTime")
                    End If
                    
                    
                   
                    '*********************************************************************************************
                    '///////////////CON ESTA CONSULTA BUSCO LA HORA DE SALIDA///////////////////////////////////
                    '*********************************************************************************************
                    Salida = "00:00"
                    Me.AdoConsulta.RecordSource = SQlSalida
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      Me.AdoConsulta.Recordset.MoveLast
                      Salida = Me.AdoConsulta.Recordset("CheckTime")
                    End If
                    
                    '*********************************************************************************************
                    '///////////////CON ESTA CONSULTA BUSCO EL NOMBRE DEL EMPLEADO///////////////////////////////////
                    '*********************************************************************************************
                    sql = "SELECT Userinfo.*, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      NombreEmpleado = Me.AdoConsulta.Recordset("Name")
                      If Not IsNull(Me.AdoConsulta.Recordset("DeptName")) Then
                       departamento = Me.AdoConsulta.Recordset("DeptName")
                      End If
                    End If
                    
                     
                    
                    '*********************************************************************************************
                    '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                    '*********************************************************************************************
                    
                    Dim Horas As String
                    
                    HorasTrabajadas = 0
                    If Salida <> "00:00" Then
                     If Entrada <> "00:00" Then
'                      HorasTrabajadas = (DateDiff("h", Entrada, Salida))
                       HorasTrabajadas = ConvertirSegundos((DateDiff("s", Entrada, Salida)), DiaInicio)
                      HoraSalida = Format(Salida, "hh:mm:ss")
                     Else
                      HorasTrabajadas = 0
                     End If
                    End If
                    
                    HorasExtras = 0
                    Horas = "0:00"
                    
                    
                        If Salida <> "00:00" Then
                         If Entrada <> "00:00" Then
                            If OutTime <> "?" Then
                              HoraSalidaHorario = OutTime
                            End If
                            
                            '***********************************************************************************
                            '//////////////VERIFICO SI LAS HORAS EXTRAS SE CALCULAN POR HORAS TRABAJADAS ///////
                            '***********************************************************************************
                            If ConfCalcularHorasTrab = False Then
                               Horas = ConvertirSegundos((DateDiff("s", HoraSalidaHorario, HoraSalida)), DiaInicio)
                            ElseIf CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) > ConfHorasTrabajadas Then
                               HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - ConfHorasTrabajadas) * 3600
                               Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                               
                            End If
                            
                         Else
                             HorasExtras = 0
                         End If
                        Else
                         HorasExtras = 0
                        End If

                    
'                    If HorasTrabajadas < 0 Then
'                      HorasTrabajadas = 0
'                    End If
                    
'                    If HorasExtras < 0 Then
'                      HorasExtras = 0
'                    End If
                    
                            Me.AdoReportes.Recordset.AddNew
                             Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                             Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                             Me.AdoReportes.Recordset("Campo3") = departamento
                             Me.AdoReportes.Recordset("CampoFecha1") = Entrada
                             If Salida <> "" Then
                               Me.AdoReportes.Recordset("CampoFecha2") = Salida
                             End If
                             Me.AdoReportes.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                             Me.AdoReportes.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
                             Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                             Me.AdoReportes.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                            Me.AdoReportes.Recordset.Update
                            
                
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      

      
     sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca From Reportes ORDER BY Reportes.CampoNum1,Reportes.CampoFecha3"
     Me.AdoExporta.RecordSource = sql
     Me.AdoExporta.Refresh


     
     
 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 '///////////////////////////////////CREO EL ARCHIVO TXT DE EXPORTACION /////////////////////////////////////////////////////////////
 '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
 
 
Me.AdoExporta.Refresh
salir = False
osProgress1.Visible = True
Me.CommonDialog1.ShowSave
Directorio = ""
Directorio = Me.CommonDialog1.FileName
AdoExporta.Recordset.MoveLast
Maximo = AdoExporta.Recordset.RecordCount
If (Dir(Directorio) <> "") Then
  Respuesta = MsgBox("Reescribir el Archivo?", vbYesNo, "Zeus Contabilidad")
  If Respuesta = 6 Then
               
               Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoExporta.Recordset.MoveFirst
                With osProgress1
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   j = 0
                 Do While Not AdoExporta.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                    

                    Cadena = Me.AdoExporta.Recordset("FechaMarca") & " " & Format(Me.AdoExporta.Recordset("Entrada"), "HH:MM") & " " & Format(Me.AdoExporta.Recordset("CodEmpleado"), "0000000000000#") & " 00"
                    Print #1, Cadena
                    Cadena = Me.AdoExporta.Recordset("FechaMarca") & " " & Format(Me.AdoExporta.Recordset("Salida"), "HH:MM") & " " & Format(Me.AdoExporta.Recordset("CodEmpleado"), "0000000000000#") & " 00"
                    Print #1, Cadena
                    
                    
                  AdoExporta.Recordset.MoveNext
                  j = j + 1
                  Me.Caption = "Procesando:  " & j & " de " & Maximo & " Registros "
                  DoEvents
                  .Value = j
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Sistema de Enlace"
                salir = True
  End If
Else '//////En caso que no exista el Archivo///////////
                
                Open Directorio For Output As #1
                'SQLExporta = "SELECT Empleado.CodEmpleado, Empleado.CodDepartamento, Historico.CodCuenta, Historico.CuentaCredito, DetalleNomina.NumNomina, Nomina.Fecha, [DetalleNomina]![SalarioBasico]+[DetalleNomina]![Destajo]+[DetalleNomina]![HorasExtras]+[DetalleNomina]![Comisiones]+[DetalleNomina]![Incentivos]-[DetalleNomina]![Deducciones]-[DetalleNomina]![Prestamo]-[DetalleNomina]![MontoINSS]-[DetalleNomina]![MontoIR]+[DetalleNomina]![TotalSubsidio] AS GranTotal FROM Nomina INNER JOIN ((Empleado INNER JOIN DetalleNomina ON Empleado.CodEmpleado = DetalleNomina.CodEmpleado) INNER JOIN Historico ON Empleado.CodEmpleado = Historico.Codempleado) ON Nomina.NumNomina = DetalleNomina.NumNomina Where DetalleNomina.NumNomina = " & NumNomina & " ORDER BY Empleado.CodEmpleado"
                
                AdoExporta.Recordset.MoveFirst
                With osProgress1
                   .Min = 0
                   .Value = 0
                   .Max = Maximo
                   j = 0
                 Do While Not AdoExporta.Recordset.EOF
                 '////////Inicialiso las variables/////////////////
                 
                    If Me.AdoExporta.Recordset("Entrada") <> "12:00:00 a.m." Then
                        Cadena = Me.AdoExporta.Recordset("FechaMarca") & " " & Format(Me.AdoExporta.Recordset("Entrada"), "HH:MM") & " " & Format(Me.AdoExporta.Recordset("CodEmpleado"), "0000000000000#") & " 00"
                        Print #1, Cadena
                    End If
                    
                    If Me.AdoExporta.Recordset("Salida") <> "12:00:00 a.m." Then
                        Cadena = Me.AdoExporta.Recordset("FechaMarca") & " " & Format(Me.AdoExporta.Recordset("Salida"), "HH:MM") & " " & Format(Me.AdoExporta.Recordset("CodEmpleado"), "0000000000000#") & " 00"
                        Print #1, Cadena
                    End If
                                    
                    
                    
                  AdoExporta.Recordset.MoveNext
                  j = j + 1
                  .Value = j
                  Me.Caption = "Procesando:  " & j & " de " & Maximo & " Registros "
                  DoEvents
                  Cadena = ""
                  Loop
                  End With
                  
                 Close #1

                MsgBox "La Exportacion, fue Creada con Exito", vbExclamation, "Zeus Facturacion"
 
 
 
 
 
 
 
 End If
 
     Me.DBGTransacciones.Columns(0).Width = 1300
     Me.DBGTransacciones.Columns(3).Width = 2200
     Me.DBGTransacciones.Columns(4).Width = 2200
End Sub
