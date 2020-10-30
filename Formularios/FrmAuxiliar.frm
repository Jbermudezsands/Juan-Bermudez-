VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmAuxiliar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarjeta Auxiliar de Marcadas"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15900
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   15900
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   6495
      Left            =   14160
      TabIndex        =   11
      Top             =   1200
      Width           =   1575
      _Version        =   786432
      _ExtentX        =   2778
      _ExtentY        =   11456
      _StockProps     =   79
      UseVisualStyle  =   -1  'True
      Begin XtremeSuiteControls.PushButton PushButton4 
         Height          =   495
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Asistencia 2"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   495
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   873
         _StockProps     =   79
         Caption         =   "Tarjeta Auxiliar"
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   6000
         Width           =   1335
         _Version        =   786432
         _ExtentX        =   2355
         _ExtentY        =   661
         _StockProps     =   79
         Caption         =   "Salir"
         UseVisualStyle  =   -1  'True
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
      ScaleWidth      =   15975
      TabIndex        =   9
      Top             =   0
      Width           =   15975
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   15960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Label lbltitulo 
         BackStyle       =   0  'Transparent
         Caption         =   "Tarjeta  Auxiliar de Marcadas"
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
         Left            =   5280
         TabIndex        =   10
         Top             =   360
         Width           =   3825
      End
      Begin VB.Image Image1 
         Height          =   1020
         Left            =   240
         Picture         =   "FrmAuxiliar.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos Generales de la Cuenta"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   13935
      Begin ACTIVESKINLibCtl.SkinLabel LblNombres 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmAuxiliar.frx":16F0
         TabIndex        =   13
         Top             =   840
         Width           =   4215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAuxiliar.frx":174E
         TabIndex        =   12
         Top             =   840
         Width           =   855
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   975
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   8535
         _Version        =   786432
         _ExtentX        =   15055
         _ExtentY        =   1720
         _StockProps     =   79
         Caption         =   "Consulta"
         UseVisualStyle  =   -1  'True
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   360
            OleObjectBlob   =   "FrmAuxiliar.frx":17B8
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin XtremeSuiteControls.CheckBox ChkTodos 
            Height          =   255
            Left            =   5880
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   2535
            _Version        =   786432
            _ExtentX        =   4471
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Consultar todos los Registros"
            UseVisualStyle  =   -1  'True
         End
         Begin XtremeSuiteControls.PushButton PushButton1 
            Height          =   375
            Left            =   5160
            TabIndex        =   7
            Top             =   360
            Width           =   1455
            _Version        =   786432
            _ExtentX        =   2566
            _ExtentY        =   661
            _StockProps     =   79
            Caption         =   "Consultar"
            UseVisualStyle  =   -1  'True
         End
         Begin MSComCtl2.DTPicker DTFechaFin 
            Height          =   285
            Left            =   3240
            TabIndex        =   6
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   77987841
            CurrentDate     =   41027
         End
         Begin MSComCtl2.DTPicker DTPFechaIni 
            Height          =   285
            Left            =   960
            TabIndex        =   5
            Top             =   360
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   503
            _Version        =   393216
            Format          =   77987841
            CurrentDate     =   41027
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   2760
            OleObjectBlob   =   "FrmAuxiliar.frx":1820
            TabIndex        =   15
            Top             =   360
            Width           =   615
         End
      End
      Begin TrueOleDBList80.TDBCombo TDBEmpleados 
         Bindings        =   "FrmAuxiliar.frx":1888
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   8811
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
         ListField       =   "Userid"
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
         _PropDict       =   $"FrmAuxiliar.frx":18A3
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
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmAuxiliar.frx":194D
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin Threed.SSCommand CmdBuscaCuenta 
         Height          =   375
         Left            =   3720
         TabIndex        =   20
         Top             =   320
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
         Picture         =   "FrmAuxiliar.frx":19BB
         Caption         =   "Buscar"
         ButtonStyle     =   4
         PictureAlignment=   9
      End
   End
   Begin TrueOleDBGrid80.TDBGrid DBGCuentas 
      Bindings        =   "FrmAuxiliar.frx":1F55
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   13935
      _ExtentX        =   24580
      _ExtentY        =   8281
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
      Splits(0).Caption=   "Tarjeta Auxiliar de Empleado"
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(6)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(9)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   3
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      PictureCurrentRow.CLSID=   "{0BE35204-8F91-11CE-9DE3-00AA004BB851}"
      PictureCurrentRow(0)=   "bHQAAOYBAABCTeYBAAAAAAAANgAAACgAAAAPAAAACQAAAAEAGAAAAAAAsAEAAAAAAAAAAAAAAAAA"
      PictureCurrentRow(1)=   "AAAAAAD///////////////////////////////////////////////////////////8AAAD/////"
      PictureCurrentRow(2)=   "//////////////////////////////////////////////////////8AAAD///////8AhgAAhgAA"
      PictureCurrentRow(3)=   "hgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgD///////8AAAD///////8AhgD///+EhoSEhoSEhoSE"
      PictureCurrentRow(4)=   "hoSEhoSEhoSEhoSEhoQAhgD///////8AAAD///////8AhgD////Gx8bGx8bGx8bGx8bGx8bGx8bG"
      PictureCurrentRow(5)=   "x8aEhoQAhgD///////8AAAD///////8AhgD///////////////////////////////////8AhgD/"
      PictureCurrentRow(6)=   "//////8AAAD///////8AhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgAAhgD///////8AAAD/"
      PictureCurrentRow(7)=   "//////////////////////////////////////////////////////////8AAAD/////////////"
      PictureCurrentRow(8)=   "//////////////////////////////////////////////8AAAA="
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
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.bgcolor=&HCE9D9D&,.bold=-1"
      _StyleDefs(20)  =   ":id=22,.fontsize=1200,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(21)  =   ":id=22,.fontname=Script MT Bold"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.bgcolor=&H8000000F&"
      _StyleDefs(23)  =   ":id=14,.fgcolor=&H0&"
      _StyleDefs(24)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(25)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(26)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(27)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(28)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(29)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(30)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(31)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(32)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(33)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(34)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(35)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(36)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(37)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(38)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(39)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(40)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(41)  =   "Named:id=33:Normal"
      _StyleDefs(42)  =   ":id=33,.parent=0"
      _StyleDefs(43)  =   "Named:id=34:Heading"
      _StyleDefs(44)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(45)  =   ":id=34,.wraptext=-1"
      _StyleDefs(46)  =   "Named:id=35:Footing"
      _StyleDefs(47)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   "Named:id=36:Selected"
      _StyleDefs(49)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(50)  =   "Named:id=37:Caption"
      _StyleDefs(51)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(52)  =   "Named:id=38:HighlightRow"
      _StyleDefs(53)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(54)  =   "Named:id=39:EvenRow"
      _StyleDefs(55)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(56)  =   "Named:id=40:OddRow"
      _StyleDefs(57)  =   ":id=40,.parent=33"
      _StyleDefs(58)  =   "Named:id=41:RecordSelector"
      _StyleDefs(59)  =   ":id=41,.parent=34"
      _StyleDefs(60)  =   "Named:id=42:FilterBar"
      _StyleDefs(61)  =   ":id=42,.parent=33"
   End
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   240
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
   Begin MSAdodcLib.Adodc AdoMarcas 
      Height          =   375
      Left            =   240
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
      Caption         =   "AdoMarcas"
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
      Left            =   3720
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
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   7320
      Visible         =   0   'False
      Width           =   13935
      _Version        =   786432
      _ExtentX        =   24580
      _ExtentY        =   661
      _StockProps     =   93
      BackColor       =   14737632
      Scrolling       =   1
      Appearance      =   6
   End
   Begin MSAdodcLib.Adodc AdoReportes 
      Height          =   375
      Left            =   7320
      Top             =   9120
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
      Left            =   11160
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
   Begin MSAdodcLib.Adodc AdoHorarioAlmuerzo 
      Height          =   375
      Left            =   3720
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
   Begin MSAdodcLib.Adodc AdoBuscaReporte 
      Height          =   375
      Left            =   7320
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
   Begin MSAdodcLib.Adodc AdoDatosEmpresa 
      Height          =   375
      Left            =   10320
      Top             =   8400
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
End
Attribute VB_Name = "FrmAuxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset
Private sql As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer

Private Sub CmdBuscaCuenta_Click()
  Quien = "Tarjeta"
  FrmConsultaReloj.Show 1
End Sub

Private Sub DBGCuentas_FilterChange()
On Error GoTo errTdbg
    'Gets called when an action is performed on the filter bar
    Dim col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
    
    'On Error GoTo errHandler
    On Error Resume Next
    Set cols = Me.DBGCuentas.Columns
    Dim c As Integer
    
    c = DBGCuentas.col
    DBGCuentas.HoldFields
    sql = Me.AdoMarcas.Recordset.Filter   'rs.Filter
    Me.AdoMarcas.Recordset.Filter = getFilter(col, cols)
    
    DBGCuentas.col = c
    DBGCuentas.EditActive = True
Exit Sub
errTdbg:
    MsgBox Err.Description
End Sub

Private Function getFilter(col As TrueOleDBGrid80.Column, cols As TrueOleDBGrid80.Columns) As String
'Creates the SQL statement in adodc1.recordset.filter
'and only filters text currently. It must be modified to
'filter other data types.
Dim tmp As String
Dim n As Integer
Dim x As Integer

For Each col In cols
    If Trim(col.FilterText) <> "" Then
        n = n + 1
        If n > 1 Then tmp = tmp & " AND "
        Select Case Me.AdoMarcas.Recordset.Fields(x).Type   'rs.Fields(x).Type
        Case adVarWChar, adVarChar: tmp = tmp & "[" & col.DataField & "] LIKE '%" & col.FilterText & "%'"
        Case adInteger, adNumeric: tmp = tmp & "[" & col.DataField & "] = " & col.FilterText
        Case adDBTimeStamp: tmp = tmp & "[" & col.DataField & "] = #" & col.FilterText & "#"
        End Select
    End If
    x = x + 1
Next col
getFilter = tmp

End Function


Private Sub Form_Load()
Dim rs As New ADODB.Recordset

 Me.DTFechaFin.Value = Now
 Me.DTPFechaIni.Value = Now

 MDIPrimero.Skin1.ApplySkin hWnd
 Me.DBGCuentas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.CmdBuscaCuenta.BackColor = RGB(222, 227, 247)
 Me.DBGCuentas.OddRowStyle.BackColor = &H80000005
 Me.DBGCuentas.AlternatingRowStyle = True
 
With Me.AdoDatosEmpresa
   .ConnectionString = ConexionReloj
   .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
   .Refresh
End With
 
With Me.AdoBuscaReporte
  .ConnectionString = ConexionReloj
End With

With Me.AdoEmpleados
  .ConnectionString = ConexionEasy
End With

With Me.AdoConsulta
  .ConnectionString = ConexionEasy
End With

With Me.AdoMarcas
  .ConnectionString = ConexionReloj
End With

With Me.AdoReportes
  .ConnectionString = ConexionReloj
End With

With Me.AdoHorarioAlmuerzo
  .ConnectionString = ConexionReloj
End With

With Me.AdoHorarios
  .ConnectionString = ConexionEasy
End With

rs.Open "DELETE FROM [Reportes] ", ConexionReloj

Me.AdoMarcas.RecordSource = "SELECT Reportes.CampoFecha1 AS Fecha, Reportes.Campo1 AS HEntrada, Reportes.Campo2 AS HEntradaComida, Reportes.Campo3 AS HSalidaComida, Reportes.Campo4 AS HSalida, Reportes.Campo5 AS MEntrada, Reportes.Campo6 AS MEntradaComida, Reportes.Campo7 AS MSalidaComida, Reportes.Campo8 AS MSalida, Reportes.Campo9 AS Laboradas, Reportes.Campo10 AS Extras FROM Reportes Where Campo1='-1'"
Me.AdoMarcas.Refresh
Me.DBGCuentas.DataSource = Me.AdoMarcas

  sql = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid ORDER BY Userinfo.Name"
Me.AdoEmpleados.RecordSource = sql
Me.AdoEmpleados.Refresh



End Sub

Private Sub PushButton1_Click()

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
Dim ContRegistros As Double, CodigoHora As String, EntradaAlmuerzo As String, SalidaAlmuerzo As String, EntradaAlmuerzo1 As String, EntradaAlmuerzo2 As String, SalidaAlmuerzo1 As String, SalidaAlmuerzo2 As String, ExcluirSabado As Boolean
Dim SQlEntradaAlmuerzo As String, SqlSalidaAlmuerzo As String, TineJornadas As Boolean
Dim EntradaA As String, SalidaA As String, DiaExtra As Double
Dim CodigoJornada As String, HorasLaborales As Double, RangoHora1 As String, RangoHora2 As String, JornadaIntercalada As Boolean, TieneJornadas As Boolean
Dim TotalTrabajadas As Date, TotalExtras As Date, HorasIn As String, SinHorario As Boolean
Dim MinutosExtra As Double, MinutosHorasExtra As Double
Dim CantHorarios As Double, SqlIN(6) As String, SqlOut(6) As String, L As Double, HoraInTime(6) As String, HoraOutTime(6) As String, MinutosTardeHorario(6) As String


Me.PushButton1.Enabled = False

      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
        If Not IsNull(Me.AdoDatosEmpresa.Recordset("MinutosExtra")) Then
         MinutosExtra = Me.AdoDatosEmpresa.Recordset("MinutosExtra")
        Else
         MinutosExtra = 0
        End If
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.Userid)='" & Me.TDBEmpleados.Text & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      
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
        
        Contador = 0
        FechaInicial = Format(Me.DTPFechaIni.Value, "dd/mm/yyyy")
        Do While FechaInicial <= DTFechaFin.Value
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
                
                '///////////CALCULO EL NUMERO DE DIAS ENTRE HORARIO Y SELECCIONADA ///////////////
                ' Diferencias en dias
                'DateDiff("d", "01/01/2000 14:39:00","01/01/2006 14:00:00")
                'Diferencias en horas
                'DateDiff("h", "01/01/2000 14:39:00","01/01/2006 14:00:00")
                'Diferencias en minutos
                'DateDiff("n", "01/01/2000 14:39:00","01/01/2006 14:00:00")
        '        Date1 = Format(CDate(FechaInicioH), "dd/mm/yyyy")
        '        Date2 = Format(CDate(Me.DtpFechaINI.Value), "dd/mm/yyyy")
        
     
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                '////////////BUSCO EL HORARIO PARA ESTE EMPLEADO ////////////////////////////////////////////////////////////////
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                 Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
                
                
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
               If Not Me.AdoHorarios.Recordset.EOF Then
                    
                     Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
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
                           
                           ConfCalcularHorasTrab = True
                           ConfHorasTrabajadas = 0
                           
                           
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                        SinHorario = True
                      Else
                                CantHorarios = 0
                                Me.AdoHorarios.Refresh
                                SinHorario = False
                                 Do While Not Me.AdoHorarios.Recordset.EOF
        
                                               CodigoHora = Me.AdoHorarios.Recordset("Schid")
                                               CodigoH = Me.AdoHorarios.Recordset("Schid")
                        
                                               TieneJornadas = False
                        
                                               '*******************************************************************************************************************
                                               '*********************************BUSCO EL HORARIO DE ALMUERZO *****************************************************
                                               '*******************************************************************************************************************
                                               Me.AdoHorarioAlmuerzo.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHora & "))"
                                               Me.AdoHorarioAlmuerzo.Refresh
                                               If Not Me.AdoHorarioAlmuerzo.Recordset.EOF Then
                                                 EntradaAlmuerzo = Me.AdoHorarioAlmuerzo.Recordset("EntradaAlmuerzo")
                                                 SalidaAlmuerzo = Me.AdoHorarioAlmuerzo.Recordset("SalidaAlmuerzo")
                                                 EntradaAlmuerzo1 = Me.AdoHorarioAlmuerzo.Recordset("EntradaAlmuerzo1")
                                                 EntradaAlmuerzo2 = Me.AdoHorarioAlmuerzo.Recordset("EntradaAlmuerzo2")
                                                 SalidaAlmuerzo1 = Me.AdoHorarioAlmuerzo.Recordset("SalidaAlmuerzo1")
                                                 SalidaAlmuerzo2 = Me.AdoHorarioAlmuerzo.Recordset("SalidaAlmuerzo2")
                                                 ExcluirSabado = Me.AdoHorarioAlmuerzo.Recordset("ExcluirSabado")
                                               End If
                        
                        
                                                '********************************************************************************************
                                                '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                                '********************************************************************************************
                        
                                                CodigoHora = Me.AdoHorarios.Recordset("Schid")
                                                Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHora & "))"
                                                Me.AdoBuscaReporte.Refresh
                                                If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                '/////SI TIENE HORAS EXTRA EN EL HORARIO, SE CAMBIA LA CONFIGURACION GENERAL ////////////
                                                DiaExtra = DiaSemana(Day(FechaInicial), Month(FechaInicial), Year(FechaInicial))
                                                TipoHorasTrabajada = Me.AdoBuscaReporte.Recordset("TipoCalcularHorasTrab")
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
                                               FechaHFinal = CDate(Format(FechaInicial, "dd/mm/yyyy") & " " & BInTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                                               
                                
                                
                        
                                               sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                        
                                               FechaEntradaInicio = FechaHInicio
                                               FechaEntradaFin = FechaHFinal
                        
'                                               FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                                               FechaHFinal = CDate(Format(FechaInicial, "dd/mm/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
'
'                                               HorasIn = Int(LongitudMinutosIn / 60) & ":" & Int(LongitudMinutosIn Mod 60)
'                                               If (CDate(InTime) + CDate(HorasIn)) > CDate("23:59") Then
'                                                '////SI LA SALIDA ES PARA EL DIA SIGUIENTE PASO PARA EL DIA SIGUIENTE
'                                                FechaHInicio = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & BOutTime & "#"
'                                                FechaHFinal = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & EOutTime & "#" '+ CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                        '                        FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
'                                               End If
                        
                        '                       SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, CheckiEout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                        '                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                               
                                                                                          ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                                LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                                HorasIn = DateAdd("n", LongitudMinutosIn, Format(FechaInicial, "mm/dd/yyyy") & " " & EInTime)
                                                FechaHInicio = "#" & Mid(HorasIn, 1, 10) & " " & BOutTime & "#" 'Me.DtpFechaINI.Value
                                                MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                                MinutosTarde = MinutosSalida & ":00" & ":00"
                                                LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                                HorasIn = DateAdd("n", LongitudMinutosIn, Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime)
'                                                FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                                FechaHFinal = "#" & Mid(Format(HorasIn, "mm/dd/yyyy"), 1, 10) & " " & EOutTime & "#"
                                                    
                                               SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText, Checkinout.CheckType FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                                                           "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                        
                                               FechaSalidaInicio = FechaHInicio
                                               FechaSalidaFin = FechaHFinal
                        
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaInicio, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       If ConfHorasTrabajadas < 0 Then
                                         FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                         ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       End If
                                    End If
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                                '///////////////////////////////BUSCO EL HORARIO DEL ALMUERZO //////////////////////////////////////////////////////////
                                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        
                                                FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & EntradaAlmuerzo1 & "#"
                                                FechaHFinal = CDate(FechaInicial)
                                                FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EntradaAlmuerzo2 & "#"
                                                SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                                                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                        
                        
                                                FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                                                FechaHFinal = CDate(FechaInicial)
                                                FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                                                SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid " & _
                                                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
'
                         
                                 SqlIN(CantHorarios) = sql
                                 SqlOut(CantHorarios) = SQlSalida
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
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName FROM (((Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid) INNER JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber)  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName FROM (((Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid) INNER JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber)  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                      End If

                    
                    
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////////////BUSCO EL HORARIO DEL ALMUERZO //////////////////////////////////////////////////////////
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        
                        '*******************************************************************************************************************
                       '*********************************BUSCO EL HORARIO DE ALMUERZO *****************************************************
                       '*******************************************************************************************************************
                       Me.AdoHorarioAlmuerzo.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.PersonalSinHorario)=True)) "
                       Me.AdoHorarioAlmuerzo.Refresh
                       If Not Me.AdoHorarioAlmuerzo.Recordset.EOF Then
                            EntradaAlmuerzo = Me.AdoHorarioAlmuerzo.Recordset("EntradaAlmuerzo")
                            SalidaAlmuerzo = Me.AdoHorarioAlmuerzo.Recordset("SalidaAlmuerzo")
                            EntradaAlmuerzo1 = Me.AdoHorarioAlmuerzo.Recordset("EntradaAlmuerzo1")
                            EntradaAlmuerzo2 = Me.AdoHorarioAlmuerzo.Recordset("EntradaAlmuerzo2")
                            SalidaAlmuerzo1 = Me.AdoHorarioAlmuerzo.Recordset("SalidaAlmuerzo1")
                            SalidaAlmuerzo2 = Me.AdoHorarioAlmuerzo.Recordset("SalidaAlmuerzo2")
                            ExcluirSabado = Me.AdoHorarioAlmuerzo.Recordset("ExcluirSabado")
                       
                            FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & EntradaAlmuerzo1 & "#"
                            FechaHFinal = CDate(FechaInicial)
                            FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EntradaAlmuerzo2 & "#"
                            SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                            
                            
                            FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                            FechaHFinal = CDate(FechaInicial)
                            FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                            SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Checkinout.CheckType = Status.Statusid  " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                            
                      End If
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                        SinHorario = True
                    
                 End If
                    
                        For L = 0 To CantHorarios - 1
                            sql = SqlIN(L)
                            SQlSalida = SqlOut(L)
                    
                                '*********************************************************************************************
                                '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA ALMUERZO///////////////////////////////////
                                '*********************************************************************************************
                                EntradaA = "00:00"
                                If SQlEntradaAlmuerzo <> "" Then
                                    Me.AdoConsulta.RecordSource = SQlEntradaAlmuerzo
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                      EntradaA = Me.AdoConsulta.Recordset("CheckTime")
                                    End If
                                End If
        
       
                                '*********************************************************************************************
                                '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA ALMUERZO///////////////////////////////////
                                '*********************************************************************************************
                                SalidaA = "00:00"
                                If SqlSalidaAlmuerzo <> "" Then
                                    Me.AdoConsulta.RecordSource = SqlSalidaAlmuerzo
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                      SalidaA = Me.AdoConsulta.Recordset("CheckTime")
                                    End If
                                End If
                                
                                If ExcluirSabado = True Then
                                  If DiaInicio = 6 Then
                                    EntradaA = "00:00"
                                    SalidaA = "00:00"
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
                                Salida = "00:00"
                                If TieneJornadas = True Then
                                   
                                     '///////////////////////////////CON ESTAS FECHAS BUSCO LA HORA DE SALIDA DE LA JORNADA ///////////////////
                                     
                                     
                                     HoraSalida = CDate(Entrada) + CDate(CInt(HorasLaborales) & ":00:00")
                                     FechaHInicio = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") - CDate(RangoHora1 & ":00")), "mm/dd/yyyy hh:mm:ss")
                                     FechaHFinal = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") + CDate(RangoHora2 & ":00")), "mm/dd/yyyy hh:mm:ss")
'                                     HoraSalida = Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59"
                                     HoraSalida = FechaHFinal
                                     HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                     If JornadaIntercalada = False Then
'                                        If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                           FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                        End If
'                                     End If
                               
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
                            If Entrada <> "00:00" Then
                              If ConfCalcularHorasTrab = True Then
                                  If TipoHorasTrabajada = "HorasTrab" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
                                  End If
                              End If
                            End If
                     
                    
                                '*********************************************************************************************
                                '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                                '*********************************************************************************************
                                
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
                                          HoraSalidaHorario = OutTime
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
                                            End If
                                        End If
                                        
                                        
                                     Else
                                         HorasExtras = 0
                                     End If
                                    Else
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
                    


                   
                                    Me.AdoReportes.Recordset.AddNew
                                     Me.AdoReportes.Recordset("CampoFecha1") = Format(FechaInicial, "dd/mm/yyyy")
                                     Me.AdoReportes.Recordset("Campo11") = CodEmpleado
                                     Me.AdoReportes.Recordset("Campo12") = NombreEmpleado
                                     Me.AdoReportes.Recordset("Campo1") = InTime
                                     Me.AdoReportes.Recordset("Campo2") = Format(EntradaAlmuerzo, "hh:mm")
                                     Me.AdoReportes.Recordset("Campo3") = Format(SalidaAlmuerzo, "hh:mm")
                                      Me.AdoReportes.Recordset("Campo4") = OutTime
                                      Me.AdoReportes.Recordset("CampoFecha5") = Format(Entrada, "hh:mm:ss")
                                      Me.AdoReportes.Recordset("CampoFecha6") = Format(EntradaA, "hh:mm:ss")
                                      Me.AdoReportes.Recordset("CampoFecha7") = Format(SalidaA, "hh:mm:ss")
                                      Me.AdoReportes.Recordset("CampoFecha8") = Format(Salida, "hh:mm:ss")
                                      Me.AdoReportes.Recordset("Campo10") = Format(HorasTrabajadas, "hh:mm")
                                      Me.AdoReportes.Recordset("Campo11") = Format(Horas, "hh:mm") 'HorasExtras
                                     Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                    Me.AdoReportes.Recordset.Update
                            
                     Next
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        
        Me.Caption = "Procesando Registro: " & Contador & " FECHA: " & Format(FechaInicial, "dd/mm/yyyy")

        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
'        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
        DoEvents
      Loop
      
      Me.AdoReportes.Refresh
      
'      Me.AdoMarcas.RecordSource = "SELECT Reportes.CampoFecha1 AS Fecha, Reportes.Campo1 AS HEntrada, Reportes.Campo2 AS HEntradaComida, Reportes.Campo3 AS HSalidaComida, Reportes.Campo4 AS HSalida, Reportes.CampoFecha5 AS MEntrada, Reportes.CampoFecha6 AS MEntradaComida, Reportes.CampoFecha7 AS MSalidaComida, Reportes.CampoFecha8 AS MSalida, Reportes.Campo10 AS Laboradas, Reportes.Campo11 AS Extras FROM Reportes Where Campo11='" & CodEmpleado & "'"
      Me.AdoMarcas.RecordSource = "SELECT Reportes.CampoFecha1 AS Fecha, Reportes.Campo1 AS HEntrada, Reportes.Campo2 AS HEntradaComida, Reportes.Campo3 AS HSalidaComida, Reportes.Campo4 AS HSalida, Reportes.CampoFecha5 AS MEntrada, Reportes.CampoFecha6 AS MEntradaComida, Reportes.CampoFecha7 AS MSalidaComida, Reportes.CampoFecha8 AS MSalida, Reportes.Campo10 AS Laboradas, Reportes.Campo11 AS Extras From Reportes Where (((Reportes.CampoNum1) = " & CodEmpleado & ")) ORDER BY Reportes.CampoFecha1"
      Me.AdoMarcas.Refresh

  Me.PushButton1.Enabled = True

End Sub

Private Sub PushButton2_Click()
  Dim sql As String, CodEmpleado As String
  Dim rpt As Object
  Dim fPreview As New FrmPreview
  
         CodEmpleado = Me.TDBEmpleados.Text

         sql = "SELECT Reportes.CampoFecha1 AS Fecha, Reportes.Campo1 AS HEntrada, Reportes.Campo2 AS HEntradaComida, Reportes.Campo3 AS HSalidaComida, Reportes.Campo4 AS HSalida, Reportes.CampoFecha5 AS MEntrada, Reportes.CampoFecha6 AS MEntradaComida, Reportes.CampoFecha7 AS MSalidaComida, Reportes.CampoFecha8 AS MSalida, Reportes.Campo10 AS Laboradas, Reportes.Campo11 AS Extras, Reportes.Campo12 AS NombreEmpleado, Reportes.CampoNum1 As CodEmpleado From Reportes Where (((Reportes.CampoNum1) = " & CodEmpleado & ")) ORDER BY Reportes.CampoFecha1 "
         Set rpt = New ArepDetalleAsistencia
         rpt.DataControl.ConnectionString = ConexionReloj
         rpt.DataControl.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
End Sub

Private Sub PushButton3_Click()
Unload Me
End Sub

Private Sub PushButton4_Click()
    Dim sql As String, CodEmpleado As String
  Dim rpt As Object
  Dim fPreview As New FrmPreview
  
         CodEmpleado = Me.TDBEmpleados.Text

         sql = "SELECT Reportes.CampoFecha1 AS Fecha, Reportes.Campo1 AS HEntrada, Reportes.Campo2 AS HEntradaComida, Reportes.Campo3 AS HSalidaComida, Reportes.Campo4 AS HSalida, Reportes.CampoFecha5 AS MEntrada, Reportes.CampoFecha6 AS MEntradaComida, Reportes.CampoFecha7 AS MSalidaComida, Reportes.CampoFecha8 AS MSalida, Reportes.Campo10 AS Laboradas, Reportes.Campo11 AS Extras, Reportes.Campo12 AS NombreEmpleado, Reportes.CampoNum1 As CodEmpleado From Reportes Where (((Reportes.CampoNum1) = " & CodEmpleado & ")) ORDER BY Reportes.CampoFecha1 "
         Set rpt = New ArepDetalleAsistencia3
         rpt.DataControl.ConnectionString = ConexionReloj
         rpt.DataControl.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
         
End Sub

Private Sub TDBEmpleados_Change()
  Me.AdoConsulta.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid Where (((Userinfo.UserID) = '" & Me.TDBEmpleados.Text & "')) ORDER BY Userinfo.Name"
  Me.AdoConsulta.Refresh
  If Not Me.AdoConsulta.Recordset.EOF Then
     Me.LblNombres.Caption = Me.AdoConsulta.Recordset("Name")
    
  End If
End Sub

Private Sub TDBEmpleados_ItemChange()
 
  Me.AdoConsulta.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid Where (((Userinfo.UserID) = '" & Me.TDBEmpleados.Text & "')) ORDER BY Userinfo.Name"
  Me.AdoConsulta.Refresh
  If Not Me.AdoConsulta.Recordset.EOF Then
     If Not IsNull(Me.AdoConsulta.Recordset("Name")) Then
        Me.LblNombres.Caption = Me.AdoConsulta.Recordset("Name")
     Else
        NombreEmpleado = ""
     End If
  
     
    
  End If
  
End Sub
