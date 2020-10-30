VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmExportarExcel 
   Caption         =   "Exportacion Formato Excel"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   10920
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   8760
      TabIndex        =   22
      Top             =   6360
      Width           =   1935
      Begin VB.OptionButton Option3 
         Caption         =   "Asistencia x Fecha"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Formato 2"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Formato 1"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Value           =   -1  'True
         Width           =   1695
      End
   End
   Begin VB.CheckBox ChkAcumulado 
      Caption         =   "Calcular Acumulado Rango de Fechas"
      Height          =   315
      Left            =   480
      TabIndex        =   13
      Top             =   8760
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CheckBox ChkTodosDptos 
      Caption         =   "Incluir Todos los Departamentos"
      Height          =   315
      Left            =   480
      TabIndex        =   12
      Top             =   9120
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Consulta de Registros"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   10575
      Begin VB.CommandButton CmdConsultar2 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   7320
         OLEDropMode     =   1  'Manual
         TabIndex        =   26
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdExportar2 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   9000
         OLEDropMode     =   1  'Manual
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.CommandButton CmdExportar 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   9000
         OLEDropMode     =   1  'Manual
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   7320
         OLEDropMode     =   1  'Manual
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmExportarExcel.frx":0000
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
      Begin MSComCtl2.DTPicker DTFechaFin 
         Height          =   330
         Left            =   4680
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
         Left            =   3600
         OleObjectBlob   =   "FrmExportarExcel.frx":0076
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
      ScaleWidth      =   11055
      TabIndex        =   0
      Top             =   0
      Width           =   11055
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   13440
         Y1              =   1080
         Y2              =   1080
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
         Left            =   3720
         TabIndex        =   1
         Top             =   360
         Width           =   3345
      End
      Begin VB.Image Image1 
         Height          =   1020
         Left            =   240
         Picture         =   "FrmExportarExcel.frx":00E6
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1335
      End
   End
   Begin TrueOleDBGrid80.TDBGrid DBGTransacciones 
      Bindings        =   "FrmExportarExcel.frx":4EEBC
      Height          =   4335
      Left            =   240
      TabIndex        =   8
      Top             =   2040
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "CodigoEmpleado"
      Columns(0).DataField=   "CodEmpleado"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Nombres"
      Columns(1).DataField=   "NombreEmpleado"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Fecha"
      Columns(2).DataField=   "Fecha"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Hora"
      Columns(3).DataField=   "Hora"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Tipo"
      Columns(4).DataField=   "Tipo"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Movimientos de Indices"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=5292"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5212"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131588"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131588"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=131588"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=131588"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
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
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=46,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=50,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
      _StyleDefs(55)  =   "Named:id=33:Normal"
      _StyleDefs(56)  =   ":id=33,.parent=0"
      _StyleDefs(57)  =   "Named:id=34:Heading"
      _StyleDefs(58)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(59)  =   ":id=34,.wraptext=-1"
      _StyleDefs(60)  =   "Named:id=35:Footing"
      _StyleDefs(61)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(62)  =   "Named:id=36:Selected"
      _StyleDefs(63)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(64)  =   "Named:id=37:Caption"
      _StyleDefs(65)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(66)  =   "Named:id=38:HighlightRow"
      _StyleDefs(67)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(68)  =   "Named:id=39:EvenRow"
      _StyleDefs(69)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(70)  =   "Named:id=40:OddRow"
      _StyleDefs(71)  =   ":id=40,.parent=33"
      _StyleDefs(72)  =   "Named:id=41:RecordSelector"
      _StyleDefs(73)  =   ":id=41,.parent=34"
      _StyleDefs(74)  =   "Named:id=42:FilterBar"
      _StyleDefs(75)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc AdoDepartamento 
      Height          =   375
      Left            =   720
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
      Left            =   720
      Top             =   9840
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
      Left            =   4320
      Top             =   9840
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
      Left            =   4440
      Top             =   9240
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
      Left            =   5760
      Top             =   9000
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
      Left            =   2280
      Top             =   9960
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
      Left            =   2280
      Top             =   9720
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
   Begin MSAdodcLib.Adodc AdoHorarioAlmuerzo 
      Height          =   375
      Left            =   6000
      Top             =   9720
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
      Caption         =   "AdoHorarioAlmuerzo"
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
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   375
      Left            =   960
      Top             =   9240
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
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   6480
      Visible         =   0   'False
      Width           =   8535
      _Version        =   786432
      _ExtentX        =   15055
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar osProgress2 
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   4215
      _Version        =   786432
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   93
      BackColor       =   14737632
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoExporta 
      Height          =   375
      Left            =   4920
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
   Begin XtremeSuiteControls.GroupBox FrameDpto 
      Height          =   735
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   5295
      _Version        =   786432
      _ExtentX        =   9340
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Departamento"
      BackColor       =   14737632
      UseVisualStyle  =   -1  'True
      Begin VB.CommandButton CmdBuscaCuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2280
         Picture         =   "FrmExportarExcel.frx":4EED5
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4800
         Picture         =   "FrmExportarExcel.frx":4F023
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   240
         Width           =   375
      End
      Begin TrueOleDBList80.TDBCombo DBDptoIni 
         Bindings        =   "FrmExportarExcel.frx":4F171
         Height          =   315
         Left            =   600
         TabIndex        =   17
         Top             =   240
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
         ListField       =   "DeptName"
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
         _PropDict       =   $"FrmExportarExcel.frx":4F18F
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
      Begin TrueOleDBList80.TDBCombo DBDptoFin 
         Bindings        =   "FrmExportarExcel.frx":4F239
         Height          =   315
         Left            =   3240
         TabIndex        =   18
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
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
         ListField       =   "DeptName"
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
         _PropDict       =   $"FrmExportarExcel.frx":4F257
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
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin"
         Height          =   255
         Left            =   2880
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin TrueOleDBGrid80.TDBGrid TdbgridRegistros 
      Bindings        =   "FrmExportarExcel.frx":4F301
      Height          =   4335
      Left            =   240
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   7646
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Id"
      Columns(0).DataField=   "Id"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Cedula"
      Columns(1).DataField=   "Cedula"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Area"
      Columns(2).DataField=   "Area"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Fecha"
      Columns(3).DataField=   "Fecha"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Entrada"
      Columns(4).DataField=   "Entrada"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "Salida"
      Columns(5).DataField=   "Salida"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   6
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   688
      Splits(0)._SavedRecordSelectors=   -1  'True
      Splits(0).Caption=   "Movimientos de Indices"
      Splits(0).DividerColor=   14215660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=6"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=131588"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=3519"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=3440"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=131588"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(11)=   "Column(2).Width=2725"
      Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=2646"
      Splits(0)._ColumnProps(14)=   "Column(2)._ColStyle=131588"
      Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(16)=   "Column(3).Width=2725"
      Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=2646"
      Splits(0)._ColumnProps(19)=   "Column(3)._ColStyle=131588"
      Splits(0)._ColumnProps(20)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(21)=   "Column(4).Width=2725"
      Splits(0)._ColumnProps(22)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(23)=   "Column(4)._WidthInPix=2646"
      Splits(0)._ColumnProps(24)=   "Column(4)._ColStyle=131588"
      Splits(0)._ColumnProps(25)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(26)=   "Column(5).Width=2725"
      Splits(0)._ColumnProps(27)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(5)._WidthInPix=2646"
      Splits(0)._ColumnProps(29)=   "Column(5)._ColStyle=131588"
      Splits(0)._ColumnProps(30)=   "Column(5).Order=6"
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
      _StyleDefs(39)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
      _StyleDefs(40)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
      _StyleDefs(41)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
      _StyleDefs(42)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
      _StyleDefs(43)  =   "Splits(0).Columns(2).Style:id=58,.parent=13"
      _StyleDefs(44)  =   "Splits(0).Columns(2).HeadingStyle:id=55,.parent=14"
      _StyleDefs(45)  =   "Splits(0).Columns(2).FooterStyle:id=56,.parent=15"
      _StyleDefs(46)  =   "Splits(0).Columns(2).EditorStyle:id=57,.parent=17"
      _StyleDefs(47)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
      _StyleDefs(48)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
      _StyleDefs(49)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
      _StyleDefs(50)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
      _StyleDefs(51)  =   "Splits(0).Columns(4).Style:id=46,.parent=13"
      _StyleDefs(52)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
      _StyleDefs(53)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
      _StyleDefs(54)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
      _StyleDefs(55)  =   "Splits(0).Columns(5).Style:id=50,.parent=13"
      _StyleDefs(56)  =   "Splits(0).Columns(5).HeadingStyle:id=47,.parent=14"
      _StyleDefs(57)  =   "Splits(0).Columns(5).FooterStyle:id=48,.parent=15"
      _StyleDefs(58)  =   "Splits(0).Columns(5).EditorStyle:id=49,.parent=17"
      _StyleDefs(59)  =   "Named:id=33:Normal"
      _StyleDefs(60)  =   ":id=33,.parent=0"
      _StyleDefs(61)  =   "Named:id=34:Heading"
      _StyleDefs(62)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   ":id=34,.wraptext=-1"
      _StyleDefs(64)  =   "Named:id=35:Footing"
      _StyleDefs(65)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(66)  =   "Named:id=36:Selected"
      _StyleDefs(67)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(68)  =   "Named:id=37:Caption"
      _StyleDefs(69)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(70)  =   "Named:id=38:HighlightRow"
      _StyleDefs(71)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(72)  =   "Named:id=39:EvenRow"
      _StyleDefs(73)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(74)  =   "Named:id=40:OddRow"
      _StyleDefs(75)  =   ":id=40,.parent=33"
      _StyleDefs(76)  =   "Named:id=41:RecordSelector"
      _StyleDefs(77)  =   ":id=41,.parent=34"
      _StyleDefs(78)  =   "Named:id=42:FilterBar"
      _StyleDefs(79)  =   ":id=42,.parent=33"
   End
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   840
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
End
Attribute VB_Name = "FrmExportarExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdConsultar_Click()
Dim sql As String, CodDptoIni As String, CodDptoFin As String, HoraTarde As String, TotalHorasTarde As Date
Dim rpt As Object, FechaIni As String, FechaFin As String, CodEmpleado As String, NombreEmpleado As String, departamento As String
Dim fPreview As New FrmPreview, i As Double, Dia As String, FechaInicioH As String, Date1 As Date, Date2 As Date
Dim cn As New ADODB.Connection, DiferenciaDias As Double, DiasCiclo As Double, Periodo As Double, DiaPeriodo As Double
Dim rs As New ADODB.Recordset, FechaActual As Date, DiasSumar As Double, FechaHorario As Date
Dim DiaInicio As Double, Ciclo As Double, BInTime As String, EInTime As String, BOutTime As String, EOutTime As String, TardePermintido As Double, InTime As String, OutTime As String
Dim Entrada As String, Salida As String, HorasTrabajadas As String, HorasExtras As Double, HoraSalida As Date, HoraSalidaHorario As Date
Dim HoraEntrada As Date, HoraHorario As Date, MinutosTarde As String, Cod As Double, FechaIn As String, FechaOut As String
Dim FechaHInicio As String, FechaHFinal As String, SQlSalida As String, j As Double, b As Double, HoraLaboradas As String
Dim TotalHorasTrabajadas As Double, TotalHorasExtras As Double, HorasTarde As Double, TotalHoras As Double, HoraHorarioSalida As Date, HoraAnticipada As Double
Dim MinutosSalida As Double, LongitudMinutosIn As Double, LongitudMinutosOut As Double
Dim FechaInicial As Date, Contador As Double, HorasMinutos As Date, ConfHorasTrabajadas As Double, ConfCalcularHorasTrab As Boolean
Dim CodigoJornada As String, HorasLaborales As Double, RangoHora1 As String, RangoHora2 As String, JornadaIntercalada As Boolean, TieneJornadas As Boolean
Dim TotalTrabajadas As String, TotalExtras As Date, HorasIn As String, DiaExtra As Double
Dim Horas As String, CodigoHorario As String, ToleranciaTarde As Boolean, TipoHorasTrabajada As String, RestarAlmuerzo As Double, SinHorario As Boolean
Dim MinutosExtra As Double, MinutosHorasExtra As Double, CantHorarios As Double, SqlIN(6) As String, SqlOut(6) As String, L As Double, HoraInTime(6) As String, HoraOutTime(6) As String, MinutosTardeHorario(6) As String
Dim CodEmpleado1 As String

      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      

      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************

        If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
           sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
        Else
           sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"

        End If

      
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
                           
                           
                          If LongitudMinutosIn < 1200 Then  'Menor a 1400  12horas
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
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
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                       MinutosTardeHorario(0) = MinutosTarde
                       HoraInTime(0) = InTime
                       HoraOutTime(0) = OutTime
                        CantHorarios = 1
                        SinHorario = True
                      Else
                        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '/////////////////////SIGNICA QUE TIENE HORARIO Y TAMBIEN TIENE ASIGIINADO PARA ESTE DIA ///////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////
                        SinHorario = False
                        CantHorarios = 0
                         Me.AdoHorarios.Refresh
                        
                         Do While Not Me.AdoHorarios.Recordset.EOF
                        
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
                                   
                                   
'                                   Me.AdoHorarios.Recordset.MoveLast
                                   
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
                       
                       SqlIN(CantHorarios) = sql
                       SqlOut(CantHorarios) = SQlSalida
                       MinutosTardeHorario(CantHorarios) = MinutosTarde
                       HoraInTime(CantHorarios) = InTime
                       HoraOutTime(CantHorarios) = OutTime
                       CantHorarios = CantHorarios + 1
                       Me.AdoHorarios.Recordset.MoveNext
                     Loop

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
                    
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                End If
                    
                    

                For L = 0 To CantHorarios - 1
                 
                            sql = SqlIN(L)
                            SQlSalida = SqlOut(L)
                            MinutosTarde = MinutosTardeHorario(L)
                            InTime = HoraInTime(L)
                            OutTime = HoraOutTime(L)
                            
                            
                            If CodEmpleado = "99220" Then
                              CodEmpleado = "99220"
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
                            
                            If Entrada = Salida Then
                               Entrada = "00:00"
                               Salida = "00:00"
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
                          Else
                            Entrada = "00:00"
                          End If
        
                            RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                            
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
                                      If OutTime <> "" Then
                                        HoraSalidaHorario = OutTime
                                      End If
                                    End If
                                    
                                    '***********************************************************************************
                                    '//////////////VERIFICO SI LAS HORAS EXTRAS SE CALCULAN POR HORAS TRABAJADAS ///////
                                    '***********************************************************************************
                                    If TieneJornadas = True Then
                                       If CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) > HorasLaborales Then
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                           Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                       End If
                                    Else
                                        If ConfCalcularHorasTrab = False Then
                                          If SinHorario = False Then
                                           HorasExtras = (CDbl(((DateDiff("s", HoraSalidaHorario, HoraSalida)) / 3600))) * 3600
                                           Horas = ConvertirSegundos((DateDiff("s", HoraSalidaHorario, HoraSalida)), DiaInicio)
                                          Else
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600))) * 3600
                                           Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                          End If
                                        ElseIf CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) > ConfHorasTrabajadas Then
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) - ConfHorasTrabajadas) * 3600
                                           Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                           'Resta o no el Almuerzo
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
                            '--------------------------------------------RESTO EL TOTAL DE HORAS EXTRAS DE LOS MINUTOS ------------------------------------------------------------
                            '--------------------------------------------------------------------------------------------------------------------------------------------------------
                            If Val(MinutosExtra) <> 0 Then
                             If IsNumeric(MinutosExtra) Then
                              MinutosHorasExtra = CDbl(MinutosExtra) / 60
                              HorasExtras = HorasExtras / 3600
                              If MinutosHorasExtra > HorasExtras Then
                                 HorasExtras = 0
                                 Horas = "00:00"
                              End If
                             
                             End If
                            End If
                            
                            '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                            '/////////////////////////////////////////////////////////////////////////////////////
                          If Me.ChkAcumulado.Value = 0 Then
                                 Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                 Me.AdoConsulta.Refresh
                                 If Not Me.AdoConsulta.Recordset.EOF Then
                                            If Not IsNull(Me.AdoConsulta.Recordset("Cardnum")) Then
                                              CodEmpleado1 = Me.AdoConsulta.Recordset("Cardnum")
                                            Else
                                              CodEmpleado1 = ""
                                            End If
                                    
                                         '////////////////////////////////////////////////////////////////////////////////////////////
                                         '////////////////////AGREGO EL REGISTRO DE ENTRADA /////////////////////////////////////////
                                         '///////////////////////////////////////////////////////////////////////////////////////////
                                          If Entrada <> "" Then
                                             If Entrada <> "00:00" Then
                                                 Me.AdoReportes.Recordset.AddNew
                                                 If CodEmpleado1 <> "" Then
                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado1
                                                 Else
                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                                 End If
                                                 Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                                 Me.AdoReportes.Recordset("Campo3") = "1"
                                                 Me.AdoReportes.Recordset("CampoFecha1") = Format(CDate(Entrada), "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")
                                                 Me.AdoReportes.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                                                 Me.AdoReportes.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
                                                 Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                                 Me.AdoReportes.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                                                Me.AdoReportes.Recordset.Update
                                             End If
                                          End If
                                          
                                          Me.AdoReportes.Refresh
                                         
                                        '///////////////////////////////////////////////////////////////////////////////////////////////
                                        '/////////////////////AGREGO EL REGISTRO DE SALIDA ////////////////////////////////////////////
                                        '//////////////////////////////////////////////////////////////////////////////////////////////
                                        If Salida <> "" Then
                                          If Salida <> "00:00" Then
                                                AdoReportes.Recordset.AddNew
                                                 If CodEmpleado1 <> "" Then
                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado1
                                                 Else
                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                                 End If
                                                 Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                                 Me.AdoReportes.Recordset("Campo3") = "2"
                                                 Me.AdoReportes.Recordset("CampoFecha1") = Format(CDate(Salida), "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")
                                                 Me.AdoReportes.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                                                 Me.AdoReportes.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
                                                 Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                                 Me.AdoReportes.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                                                Me.AdoReportes.Recordset.Update
                                         End If
                                       End If
                                       
                                       Me.AdoReportes.Refresh
                                         
                                End If
                                
                                
                            ElseIf Me.ChkAcumulado.Value = 1 Then
                                 Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                 Me.AdoConsulta.Refresh
                                 If Not Me.AdoConsulta.Recordset.EOF Then
                                      Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes WHERE (((Reportes.Campo1)='" & CodEmpleado & "'))"
                                      Me.AdoBuscaReporte.Refresh
                                      If Me.AdoBuscaReporte.Recordset.EOF Then
                                         Me.AdoBuscaReporte.Recordset.AddNew
                                          Me.AdoBuscaReporte.Recordset("Campo1") = CodEmpleado
                                          Me.AdoBuscaReporte.Recordset("Campo2") = NombreEmpleado
                                          Me.AdoBuscaReporte.Recordset("Campo3") = departamento
                                          Me.AdoBuscaReporte.Recordset("CampoFecha1") = Format(FechaInicial, "dd/mm/yyyy")  'Entrada
                                          If Salida <> "" Then
                                            Me.AdoBuscaReporte.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")  'Salida
                                          End If
                                          Me.AdoBuscaReporte.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                                          Me.AdoBuscaReporte.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
                                          Me.AdoBuscaReporte.Recordset("CampoNum1") = CodEmpleado
                                          Me.AdoBuscaReporte.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                                         Me.AdoBuscaReporte.Recordset.Update
                                      Else
                                         
                                          If Salida <> "" Then
                                            Me.AdoBuscaReporte.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")  'Salida
                                          End If
                                          HorasTrabajadas = sumaHoras(HorasTrabajadas, Me.AdoBuscaReporte.Recordset("Campo4"))
                                          Horas = sumaHoras(Horas, Me.AdoBuscaReporte.Recordset("Campo5"))
                                          Me.AdoBuscaReporte.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                                          Me.AdoBuscaReporte.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
        '                                  Me.AdoBuscaReporte.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                                         Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                End If
                            
                            End If
                
                Next
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
'        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop '///////////////////CON ESTE CICLO RECORRO TODOS LOS EMPLEADOS SELECCIONADOS /////////////
      
         Me.AdoReportes.Refresh

      
         
         
      sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Tipo, Reportes.CampoFecha1 AS Hora, Reportes.CampoFecha2 AS Fecha, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca From Reportes ORDER BY Reportes.CampoFecha3, Reportes.CampoNum1,Reportes.CampoFecha1"
      Me.AdoExporta.RecordSource = sql
      Me.AdoExporta.Refresh
      
      Me.DBGTransacciones.DataSource = Me.AdoExporta

End Sub

Private Sub Command2_Click()

End Sub

Private Sub CmdConsultar2_Click()
Dim sql As String, CodDptoIni As String, CodDptoFin As String, HoraTarde As String, TotalHorasTarde As Date
Dim rpt As Object, FechaIni As String, FechaFin As String, CodEmpleado As String, NombreEmpleado As String, departamento As String
Dim fPreview As New FrmPreview, i As Double, Dia As String, FechaInicioH As String, Date1 As Date, Date2 As Date
Dim cn As New ADODB.Connection, DiferenciaDias As Double, DiasCiclo As Double, Periodo As Double, DiaPeriodo As Double
Dim rs As New ADODB.Recordset, FechaActual As Date, DiasSumar As Double, FechaHorario As Date
Dim DiaInicio As Double, Ciclo As Double, BInTime As String, EInTime As String, BOutTime As String, EOutTime As String, TardePermintido As Double, InTime As String, OutTime As String
Dim Entrada As String, Salida As String, HorasTrabajadas As String, HorasExtras As Double, HoraSalida As Date, HoraSalidaHorario As Date
Dim HoraEntrada As Date, HoraHorario As Date, MinutosTarde As String, Cod As Double, FechaIn As String, FechaOut As String
Dim FechaHInicio As String, FechaHFinal As String, SQlSalida As String, j As Double, b As Double, HoraLaboradas As String
Dim TotalHorasTrabajadas As Double, TotalHorasExtras As Double, HorasTarde As Double, TotalHoras As Double, HoraHorarioSalida As Date, HoraAnticipada As Double
Dim MinutosSalida As Double, LongitudMinutosIn As Double, LongitudMinutosOut As Double
Dim FechaInicial As Date, Contador As Double, HorasMinutos As Date, ConfHorasTrabajadas As Double, ConfCalcularHorasTrab As Boolean
Dim CodigoJornada As String, HorasLaborales As Double, RangoHora1 As String, RangoHora2 As String, JornadaIntercalada As Boolean, TieneJornadas As Boolean
Dim TotalTrabajadas As String, TotalExtras As Date, HorasIn As String, DiaExtra As Double
Dim Horas As String, CodigoHorario As String, ToleranciaTarde As Boolean, TipoHorasTrabajada As String, RestarAlmuerzo As Double, SinHorario As Boolean
Dim MinutosExtra As Double, MinutosHorasExtra As Double, CantHorarios As Double, SqlIN(6) As String, SqlOut(6) As String, L As Double, HoraInTime(6) As String, HoraOutTime(6) As String, MinutosTardeHorario(6) As String
Dim CodEmpleado1 As String, IDCard As String

      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      

      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************

        If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
           sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
        Else
           sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"

        End If

      
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
                           
                           
                          If LongitudMinutosIn < 1200 Then  'Menor a 1400  12horas
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
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
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                       MinutosTardeHorario(0) = MinutosTarde
                       HoraInTime(0) = InTime
                       HoraOutTime(0) = OutTime
                        CantHorarios = 1
                        SinHorario = True
                      Else
                        '////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '/////////////////////SIGNICA QUE TIENE HORARIO Y TAMBIEN TIENE ASIGIINADO PARA ESTE DIA ///////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////
                        SinHorario = False
                        CantHorarios = 0
                         Me.AdoHorarios.Refresh
                        
                         Do While Not Me.AdoHorarios.Recordset.EOF
                        
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
                                   
                                   
'                                   Me.AdoHorarios.Recordset.MoveLast
                                   
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
                       
                       SqlIN(CantHorarios) = sql
                       SqlOut(CantHorarios) = SQlSalida
                       MinutosTardeHorario(CantHorarios) = MinutosTarde
                       HoraInTime(CantHorarios) = InTime
                       HoraOutTime(CantHorarios) = OutTime
                       CantHorarios = CantHorarios + 1
                       Me.AdoHorarios.Recordset.MoveNext
                     Loop

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
                    
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                End If
                    
                    

                For L = 0 To CantHorarios - 1
                 
                            sql = SqlIN(L)
                            SQlSalida = SqlOut(L)
                            MinutosTarde = MinutosTardeHorario(L)
                            InTime = HoraInTime(L)
                            OutTime = HoraOutTime(L)
                            
                            
                            If CodEmpleado = "99220" Then
                              CodEmpleado = "99220"
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
                            
                            If Entrada = Salida Then
                               Entrada = "00:00"
                               Salida = "00:00"
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
                          Else
                            Entrada = "00:00"
                          End If
        
                            RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                            
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
                                      If OutTime <> "" Then
                                        HoraSalidaHorario = OutTime
                                      End If
                                    End If
                                    
                                    '***********************************************************************************
                                    '//////////////VERIFICO SI LAS HORAS EXTRAS SE CALCULAN POR HORAS TRABAJADAS ///////
                                    '***********************************************************************************
                                    If TieneJornadas = True Then
                                       If CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) > HorasLaborales Then
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                           Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                       End If
                                    Else
                                        If ConfCalcularHorasTrab = False Then
                                          If SinHorario = False Then
                                           HorasExtras = (CDbl(((DateDiff("s", HoraSalidaHorario, HoraSalida)) / 3600))) * 3600
                                           Horas = ConvertirSegundos((DateDiff("s", HoraSalidaHorario, HoraSalida)), DiaInicio)
                                          Else
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600))) * 3600
                                           Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                          End If
                                        ElseIf CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) > ConfHorasTrabajadas Then
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) - ConfHorasTrabajadas) * 3600
                                           Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                           'Resta o no el Almuerzo
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
                            '--------------------------------------------RESTO EL TOTAL DE HORAS EXTRAS DE LOS MINUTOS ------------------------------------------------------------
                            '--------------------------------------------------------------------------------------------------------------------------------------------------------
                            If Val(MinutosExtra) <> 0 Then
                             If IsNumeric(MinutosExtra) Then
                              MinutosHorasExtra = CDbl(MinutosExtra) / 60
                              HorasExtras = HorasExtras / 3600
                              If MinutosHorasExtra > HorasExtras Then
                                 HorasExtras = 0
                                 Horas = "00:00"
                              End If
                             
                             End If
                            End If
                            
                            '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                            '/////////////////////////////////////////////////////////////////////////////////////
                          If Me.ChkAcumulado.Value = 0 Then
                                 Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                 Me.AdoConsulta.Refresh
                                 If Not Me.AdoConsulta.Recordset.EOF Then
                                            If Not IsNull(Me.AdoConsulta.Recordset("Cardnum")) Then
                                              CodEmpleado1 = Me.AdoConsulta.Recordset("Cardnum")
                                            Else
                                              CodEmpleado1 = ""
                                            End If
                                            
                                    If Not IsNull(Me.AdoConsulta.Recordset("IDCard")) Then
                                            
                                         IDCard = Me.AdoConsulta.Recordset("IDCard")
                                    
                                         '////////////////////////////////////////////////////////////////////////////////////////////
                                         '////////////////////AGREGO EL REGISTRO DE ENTRADA /////////////////////////////////////////
                                         '///////////////////////////////////////////////////////////////////////////////////////////
                                          If Entrada <> "" Then
                                             If Entrada <> "00:00" Then
                                                 Me.AdoReportes.Recordset.AddNew
                                                 If CodEmpleado1 <> "" Then
                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado1
                                                 Else
                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                                 End If
                                                 Me.AdoReportes.Recordset("Campo2") = NombreEmpleado 'IDCard
                                                 Me.AdoReportes.Recordset("Campo3") = departamento
                                                 Me.AdoReportes.Recordset("CampoFecha1") = Format(CDate(Entrada), "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")
                                                 Me.AdoReportes.Recordset("CampoFecha3") = Format(CDate(Salida), "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                                                 Me.AdoReportes.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
                                                 Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                                 Me.AdoReportes.Recordset.Update
                                             End If
                                          End If
                                          
'                                          sql = "SELECT Reportes.Campo1 AS Id, Reportes.Campo2 AS Cedula, Reportes.Campo3 AS Area, Reportes.CampoFecha2 AS Fecha, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha3 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca From Reportes ORDER BY Reportes.CampoFecha3, Reportes.CampoNum1,Reportes.CampoFecha1"
                                          
                                          Me.AdoReportes.Refresh
                                         
                                        '///////////////////////////////////////////////////////////////////////////////////////////////
                                        '/////////////////////AGREGO EL REGISTRO DE SALIDA ////////////////////////////////////////////
                                        '//////////////////////////////////////////////////////////////////////////////////////////////
'                                        If Salida <> "" Then
'                                          If Salida <> "00:00" Then
'                                                AdoReportes.Recordset.AddNew
'                                                 If CodEmpleado1 <> "" Then
'                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado1
'                                                 Else
'                                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado
'                                                 End If
'                                                 Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
'                                                 Me.AdoReportes.Recordset("Campo3") = "2"
'                                                 Me.AdoReportes.Recordset("CampoFecha1") = Format(CDate(Salida), "hh:mm:ss")
'                                                 Me.AdoReportes.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")
'                                                 Me.AdoReportes.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
'                                                 Me.AdoReportes.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
'                                                 Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
'                                                 Me.AdoReportes.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
'                                                Me.AdoReportes.Recordset.Update
'                                         End If
'                                       End If
                                       
                                       Me.AdoReportes.Refresh
                                       
                                 Else
                                   MsgBox "El IdCard en en el empleado " & CodEmpleado & " Es Nulo"
                                   
                                 End If
                                         
                                End If
                                
                                
                            ElseIf Me.ChkAcumulado.Value = 1 Then
                                 Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                 Me.AdoConsulta.Refresh
                                 If Not Me.AdoConsulta.Recordset.EOF Then
                                      Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes WHERE (((Reportes.Campo1)='" & CodEmpleado & "'))"
                                      Me.AdoBuscaReporte.Refresh
                                      If Me.AdoBuscaReporte.Recordset.EOF Then
                                         Me.AdoBuscaReporte.Recordset.AddNew
                                          Me.AdoBuscaReporte.Recordset("Campo1") = CodEmpleado
                                          Me.AdoBuscaReporte.Recordset("Campo2") = NombreEmpleado
                                          Me.AdoBuscaReporte.Recordset("Campo3") = departamento
                                          Me.AdoBuscaReporte.Recordset("CampoFecha1") = Format(FechaInicial, "dd/mm/yyyy")  'Entrada
                                          If Salida <> "" Then
                                            Me.AdoBuscaReporte.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")  'Salida
                                          End If
                                          Me.AdoBuscaReporte.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                                          Me.AdoBuscaReporte.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
                                          Me.AdoBuscaReporte.Recordset("CampoNum1") = CodEmpleado
                                          Me.AdoBuscaReporte.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                                         Me.AdoBuscaReporte.Recordset.Update
                                      Else
                                         
                                          If Salida <> "" Then
                                            Me.AdoBuscaReporte.Recordset("CampoFecha2") = Format(FechaInicial, "dd/mm/yyyy")  'Salida
                                          End If
                                          HorasTrabajadas = sumaHoras(HorasTrabajadas, Me.AdoBuscaReporte.Recordset("Campo4"))
                                          Horas = sumaHoras(Horas, Me.AdoBuscaReporte.Recordset("Campo5"))
                                          Me.AdoBuscaReporte.Recordset("Campo4") = Format(HorasTrabajadas, "hh:mm")
                                          Me.AdoBuscaReporte.Recordset("Campo5") = Format(Horas, "hh:mm") 'HorasExtras
        '                                  Me.AdoBuscaReporte.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                                         Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                End If
                            
                            End If
                
                Next
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
'        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop '///////////////////CON ESTE CICLO RECORRO TODOS LOS EMPLEADOS SELECCIONADOS /////////////
      
         Me.AdoReportes.Refresh

      
         
         
      sql = "SELECT Reportes.Campo1 AS Id, Reportes.Campo2 AS Cedula, Reportes.Campo3 AS Area, Reportes.CampoFecha2 AS Fecha, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha3 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca From Reportes ORDER BY Reportes.CampoFecha2, Reportes.CampoNum1"
      Me.AdoExporta.RecordSource = sql
      Me.AdoExporta.Refresh
      
      Me.DBGTransacciones.DataSource = Me.AdoExporta

End Sub

Private Sub CmdExportar_Click()
Dim SQlReportes As String, V As Integer, H As Integer, i As Integer
Dim Fecha As Date, Hora As String, Tipo As String, CodigoEmpleado As String

    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 2
H = 0
i = 1
           objExcel.ActiveSheet.Columns("B").NumberFormat = "@"
           objExcel.ActiveSheet.Name = "Marcas"
           objExcel.ActiveSheet.Cells(1, 1) = "CodigoEmpleado"
           objExcel.ActiveSheet.Cells(1, 2) = "Fecha"
           objExcel.ActiveSheet.Cells(1, 3) = "Hora"
           objExcel.ActiveSheet.Cells(1, 4) = "Tipo"
           
     Me.AdoExporta.Refresh
 
     Do While Not Me.AdoExporta.Recordset.EOF 'esto nos sirve pa leer los datos desde
     DoEvents
     Fecha = Me.AdoExporta.Recordset("Fecha")
     Hora = Format(CDate(Me.AdoExporta.Recordset("Hora")), "hh:mm:ss")
     CodigoEmpleado = Me.AdoExporta.Recordset("CodEmpleado")
     Tipo = Me.AdoExporta.Recordset("Tipo")

       With AdoExporta.Recordset
            'objExcel.Cells(1, 1).Format = Text
            objExcel.ActiveSheet.Cells(V, H + 1) = CodigoEmpleado
            objExcel.ActiveSheet.Cells(V, H + 2) = Format(Fecha, "dd/mm/yyyy")
            objExcel.ActiveSheet.Cells(V, H + 3) = Hora
            objExcel.ActiveSheet.Cells(V, H + 4) = Tipo
            V = V + 1
            i = i + 1
            .MoveNext
       End With
     Loop
     
     
     

       objExcel.ActiveSheet.Columns("A").ColumnWidth = 16
       objExcel.ActiveSheet.Columns("B").ColumnWidth = 11
       objExcel.ActiveSheet.Columns("C").ColumnWidth = 9
       objExcel.ActiveSheet.Columns("D").ColumnWidth = 8

         
 
 Set objExcel = Nothing 'una vez hemos terminado descargamos el objeto
End Sub

Private Sub CmdExportar2_Click()
Dim SQlReportes As String, V As Integer, H As Integer, i As Integer
Dim Fecha As Date, Entrada As String, Salida As String, Cedula As String, Id As String, Area As String
Dim HoraLaborada As String

    Call Inicio_Excel 'Llamamos a la funcion que abre el workbook en excel
    'Call Formato_Excel(8, Heading()) 'llamamos a la funcion que da el formato al nuevo workbook
    
V = 2
H = 0
i = 1
           objExcel.ActiveSheet.Columns("B").NumberFormat = "@"
           objExcel.ActiveSheet.Columns("D").NumberFormat = "@"
           objExcel.ActiveSheet.Name = "Marcas"
           objExcel.ActiveSheet.Cells(1, 1) = "Id"
           objExcel.ActiveSheet.Cells(1, 2) = "Cedula"
           objExcel.ActiveSheet.Cells(1, 3) = "Area"
           objExcel.ActiveSheet.Cells(1, 4) = "Fecha"
           objExcel.ActiveSheet.Cells(1, 5) = "Entrada"
           objExcel.ActiveSheet.Cells(1, 6) = "Salida"
           objExcel.ActiveSheet.Cells(1, 7) = "Hlaboradas"
           
     Me.AdoExporta.Refresh
 
     Do While Not Me.AdoExporta.Recordset.EOF 'esto nos sirve pa leer los datos desde
     DoEvents
     Fecha = Format(Me.AdoExporta.Recordset("Fecha"), "dd/mm/yyyy")
     Entrada = Format(CDate(Me.AdoExporta.Recordset("Entrada")), "hh:mm:ss")
     Salida = Format(CDate(Me.AdoExporta.Recordset("Salida")), "hh:mm:ss")
     Id = Me.AdoExporta.Recordset("Id")
     Area = Me.AdoExporta.Recordset("Area")
     Cedula = Me.AdoExporta.Recordset("Cedula")
     HoraLaborada = Me.AdoExporta.Recordset("HorasTrabajadas")

       With AdoExporta.Recordset
            'objExcel.Cells(1, 1).Format = Text
            objExcel.ActiveSheet.Cells(V, H + 1) = Id
            objExcel.ActiveSheet.Cells(V, H + 2) = Cedula
            objExcel.ActiveSheet.Cells(V, H + 3) = Area
            objExcel.ActiveSheet.Cells(V, H + 4) = Format(Fecha, "dd/mm/yyyy")
            objExcel.ActiveSheet.Cells(V, H + 5) = Entrada
            objExcel.ActiveSheet.Cells(V, H + 6) = Salida
            objExcel.ActiveSheet.Cells(V, H + 7) = HoraLaborada
            V = V + 1
            i = i + 1
            .MoveNext
       End With
     Loop
     
     
     

       objExcel.ActiveSheet.Columns("A").ColumnWidth = 6
       objExcel.ActiveSheet.Columns("B").ColumnWidth = 18
       objExcel.ActiveSheet.Columns("C").ColumnWidth = 20
       objExcel.ActiveSheet.Columns("D").ColumnWidth = 10

         
 
 Set objExcel = Nothing 'una vez hemos
End Sub

Private Sub Form_Load()
 MDIPrimero.Skin1.ApplySkin hWnd
 Me.DBGTransacciones.EvenRowStyle.BackColor = RGB(216, 228, 248)
' Me.CmdBuscaCuenta.BackColor = RGB(222, 227, 247)
 Me.DBGTransacciones.OddRowStyle.BackColor = &H80000005
 Me.DBGTransacciones.AlternatingRowStyle = True
 
 Me.DTFechaFin.Value = Format(Now, "dd/mm/yyyy")
Me.DTPFechaIni.Value = Format(Now, "dd/mm/yyyy")

With Me.AdoDatosEmpresa
   .ConnectionString = ConexionReloj
   .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
   .Refresh
End With

With Me.AdoExporta
  .ConnectionString = ConexionReloj
End With

With Me.AdoHorarioAlmuerzo
  .ConnectionString = ConexionReloj
End With

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


Me.AdoEmpleados2.RecordSource = "SELECT Userinfo.* FROM Userinfo"
Me.AdoEmpleados2.Refresh


Me.AdoDepartamento.RecordSource = "SELECT Dept.Deptid, Dept.DeptName FROM Dept"
Me.AdoDepartamento.Refresh
End Sub

Private Sub Option1_Click()
  If Me.Option1.Value = True Then
     Me.CmdExportar.Visible = True
     Me.CmdConsultar.Visible = True
     Me.DBGTransacciones.Visible = True
     
     
     Me.CmdExportar2.Visible = False
     Me.CmdConsultar2.Visible = False
     Me.TdbgridRegistros.Visible = False
  Else

  End If
  



End Sub

Private Sub Option2_Click()
  If Me.Option2.Value = True Then
     Me.CmdExportar.Visible = False
     Me.CmdConsultar.Visible = False
     Me.DBGTransacciones.Visible = False
     
     Me.CmdConsultar2.Visible = True
     Me.CmdExportar2.Visible = True
     Me.TdbgridRegistros.Visible = True
  Else

  End If
End Sub

Private Sub Option3_Click()
  If Me.Option3.Value = True Then
     Me.CmdExportar.Visible = True
     Me.CmdConsultar.Visible = True
     Me.DBGTransacciones.Visible = True
     
     
     Me.CmdExportar2.Visible = False
     Me.CmdConsultar2.Visible = False
     Me.TdbgridRegistros.Visible = False
  End If
End Sub
