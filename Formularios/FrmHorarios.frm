VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{065E6FD1-1BF9-11D2-BAE8-00104B9E0792}#3.0#0"; "ssa3d30.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmHorarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horarios "
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7770
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
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CONFIGURARCION DE HORARIOS"
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
         Left            =   1920
         TabIndex        =   1
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
         Picture         =   "FrmHorarios.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1170
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox3 
      Height          =   4095
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   4575
      _Version        =   786432
      _ExtentX        =   8070
      _ExtentY        =   7223
      _StockProps     =   79
      Caption         =   "Horarios Asignados - Almuerzos"
      BackColor       =   7592959
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHorarios.frx":413F
         TabIndex        =   20
         Top             =   360
         Width           =   1335
      End
      Begin TrueOleDBList80.TDBCombo TDBMatutino 
         Bindings        =   "FrmHorarios.frx":41BF
         Height          =   315
         Left            =   1680
         TabIndex        =   3
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
         _PropDict       =   $"FrmHorarios.frx":41D9
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
         TabIndex        =   4
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
      Begin XtremeSuiteControls.CheckBox ChkAsignarHorario 
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1320
         Width           =   3615
         _Version        =   786432
         _ExtentX        =   6376
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Asignar Almuerzo al Personal sin Horario"
         BackColor       =   7592959
         UseVisualStyle  =   -1  'True
      End
      Begin XtremeSuiteControls.GroupBox GroupBox1 
         Height          =   2295
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   4335
         _Version        =   786432
         _ExtentX        =   7646
         _ExtentY        =   4048
         _StockProps     =   79
         Caption         =   "Configuracion Calculo de Horas Extras"
         BackColor       =   7592959
         Begin VB.CheckBox ChkHorasExtraLaborada 
            BackColor       =   &H0073DBFF&
            Caption         =   "Calcular Hora Trab con Hora Entrada"
            Height          =   495
            Left            =   2520
            TabIndex        =   35
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox TxtHorasDomingo 
            Height          =   285
            Left            =   1440
            TabIndex        =   33
            Text            =   "0"
            Top             =   1440
            Width           =   735
         End
         Begin VB.TextBox TxtHorasSabado 
            Height          =   285
            Left            =   1440
            TabIndex        =   32
            Text            =   "0"
            Top             =   960
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblHora 
            Height          =   375
            Left            =   240
            OleObjectBlob   =   "FrmHorarios.frx":4283
            TabIndex        =   28
            Top             =   960
            Width           =   1095
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblHorasTrab 
            Height          =   495
            Left            =   240
            OleObjectBlob   =   "FrmHorarios.frx":4313
            TabIndex        =   27
            Top             =   360
            Width           =   1095
         End
         Begin VB.TextBox TxtHorasTrab 
            Height          =   285
            Left            =   1440
            TabIndex        =   7
            Text            =   "0"
            Top             =   480
            Width           =   735
         End
         Begin XtremeSuiteControls.RadioButton OptHoras 
            Height          =   495
            Left            =   2280
            TabIndex        =   8
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
            TabIndex        =   9
            Top             =   1200
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Calcular Por Hora Salida"
            BackColor       =   7592959
            UseVisualStyle  =   -1  'True
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblHoraDomingo 
            Height          =   375
            Left            =   240
            OleObjectBlob   =   "FrmHorarios.frx":43AB
            TabIndex        =   34
            Top             =   1440
            Width           =   1095
         End
         Begin XtremeSuiteControls.RadioButton OptHorario 
            Height          =   495
            Left            =   2280
            TabIndex        =   36
            Top             =   1680
            Width           =   1575
            _Version        =   786432
            _ExtentX        =   2778
            _ExtentY        =   873
            _StockProps     =   79
            Caption         =   "Calcular Segun Horario"
            BackColor       =   7592959
            UseVisualStyle  =   -1  'True
         End
      End
      Begin XtremeSuiteControls.CheckBox OptRestarAlmuerzo 
         Height          =   255
         Left            =   360
         TabIndex        =   31
         Top             =   1000
         Width           =   3615
         _Version        =   786432
         _ExtentX        =   6376
         _ExtentY        =   450
         _StockProps     =   79
         Caption         =   "Restar Periodo de Almuerzo"
         BackColor       =   7592959
         UseVisualStyle  =   -1  'True
      End
   End
   Begin XtremeSuiteControls.GroupBox GroupBox2 
      Height          =   4095
      Left            =   4800
      TabIndex        =   10
      Top             =   1200
      Width           =   2895
      _Version        =   786432
      _ExtentX        =   5106
      _ExtentY        =   7223
      _StockProps     =   79
      Caption         =   "Periodo de  Horas Almuerzo"
      BackColor       =   7592959
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHorarios.frx":443B
         TabIndex        =   30
         Top             =   2160
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel8 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHorarios.frx":44B5
         TabIndex        =   29
         Top             =   1080
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHorarios.frx":4531
         TabIndex        =   26
         Top             =   2880
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHorarios.frx":45B3
         TabIndex        =   25
         Top             =   2520
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHorarios.frx":4631
         TabIndex        =   24
         Top             =   1800
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmHorarios.frx":46B3
         TabIndex        =   23
         Top             =   1440
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHorarios.frx":4731
         TabIndex        =   22
         Top             =   720
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmHorarios.frx":47AD
         TabIndex        =   21
         Top             =   360
         Width           =   1335
      End
      Begin MSMask.MaskEdBox TxtSalida 
         Height          =   285
         Left            =   1680
         TabIndex        =   11
         Top             =   720
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
         TabIndex        =   12
         Top             =   360
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
         TabIndex        =   13
         Top             =   1800
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
         TabIndex        =   14
         Top             =   1440
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
         TabIndex        =   15
         Top             =   2880
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
         TabIndex        =   16
         Top             =   2520
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   503
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
   End
   Begin Threed.SSCommand CmdCerrar 
      Height          =   585
      Left            =   6600
      TabIndex        =   17
      ToolTipText     =   "Cerrar la ventana de nuevo Crédito"
      Top             =   5400
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
      Picture         =   "FrmHorarios.frx":482B
      Caption         =   "Cerrar"
      ButtonStyle     =   3
      PictureAlignment=   9
      PictureDnFrames =   1
      PictureDn       =   "FrmHorarios.frx":5505
   End
   Begin Threed.SSCommand CmdGuardar 
      Height          =   585
      Left            =   5400
      TabIndex        =   18
      ToolTipText     =   "Guarda el Crédito"
      Top             =   5400
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
      Picture         =   "FrmHorarios.frx":5767
      Caption         =   "Guardar  "
      ButtonStyle     =   3
      PictureAlignment=   9
      ShapeSize       =   1
   End
   Begin Threed.SSCommand CmdQuitarCredito 
      Height          =   585
      Left            =   4200
      TabIndex        =   19
      ToolTipText     =   "Quita el Crédito de la lista de pagos"
      Top             =   5400
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
      Picture         =   "FrmHorarios.frx":6041
      Caption         =   "Borrar"
      ButtonStyle     =   3
      PictureAlignment=   9
      ShapeSize       =   1
   End
   Begin MSAdodcLib.Adodc AdoHorarios 
      Height          =   375
      Left            =   720
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
      Left            =   600
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
      Top             =   7920
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
End
Attribute VB_Name = "FrmHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CmdCerrar_Click()
Unload Me
End Sub




Private Sub CmdGuardar_Click()
  Dim Codigo As Double
  Codigo = Me.TDBMatutino.Columns(0).Text
  Me.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & Codigo & "))"
  Me.AdoConsulta.Refresh
  If Me.AdoConsulta.Recordset.EOF Then
        Me.AdoConsulta.Refresh
        Me.AdoConsulta.Recordset.AddNew
        Me.AdoConsulta.Recordset("Schid") = Me.TDBMatutino.Columns(0).Text
        Me.AdoConsulta.Recordset("EntradaAlmuerzo") = Me.TxtEntrada.Text
        Me.AdoConsulta.Recordset("SalidaAlmuerzo") = Me.TxtSalida.Text
        Me.AdoConsulta.Recordset("EntradaAlmuerzo1") = Me.TxtEntrada1.Text
        Me.AdoConsulta.Recordset("EntradaAlmuerzo2") = Me.TxtEntrada2.Text
        Me.AdoConsulta.Recordset("SalidaAlmuerzo1") = Me.TxtSalida1.Text
        Me.AdoConsulta.Recordset("SalidaAlmuerzo2") = Me.TxtSalida2.Text
        If Me.ChkUnicoHorario.Value = xtpChecked Then
         Me.AdoConsulta.Recordset("ExcluirSabado") = True
        Else
         Me.AdoConsulta.Recordset("ExcluirSabado") = False
        End If
        
        If Me.ChkAsignarHorario.Value = xtpChecked Then
          Me.AdoConsulta.Recordset("PersonalSinHorario") = True
        Else
          Me.AdoConsulta.Recordset("PersonalSinHorario") = False
        End If
        
        If Me.ChkHorasExtraLaborada.Value = xtpChecked Then
          Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "HorasTrab"
        Else
          Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "N"
        End If
        
        If Me.OptHorario.Value = True Then
           Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "Horario"
        End If
        
         If Me.OptHoras.Value = False Then
           Me.AdoConsulta.Recordset("CalcularHorasTrab") = 0
         Else
           Me.AdoConsulta.Recordset("CalcularHorasTrab") = 1
        End If
        
        If Me.TxtHorasTrab.Text <> "" Then
         Me.AdoConsulta.Recordset("HorasTrab") = Me.TxtHorasTrab.Text
        End If
        
        If Me.TxtHorasSabado.Text <> "" Then
         Me.AdoConsulta.Recordset("HorasTrabSab") = Me.TxtHorasSabado.Text
        End If
        
        If Me.TxtHorasDomingo.Text <> "" Then
         Me.AdoConsulta.Recordset("HorasTrabDom") = Me.TxtHorasDomingo.Text
        End If
        
        If Me.OptRestarAlmuerzo.Value = False Then
           Me.AdoConsulta.Recordset("RestarAlmuerzo") = 0
         Else
           Me.AdoConsulta.Recordset("RestarAlmuerzo") = 1
        End If
        
        Me.AdoConsulta.Recordset.Update
   Else
        Me.AdoConsulta.Recordset("EntradaAlmuerzo") = Me.TxtEntrada.Text
        Me.AdoConsulta.Recordset("SalidaAlmuerzo") = Me.TxtSalida.Text
        Me.AdoConsulta.Recordset("EntradaAlmuerzo1") = Me.TxtEntrada1.Text
        Me.AdoConsulta.Recordset("EntradaAlmuerzo2") = Me.TxtEntrada2.Text
        Me.AdoConsulta.Recordset("SalidaAlmuerzo1") = Me.TxtSalida1.Text
        Me.AdoConsulta.Recordset("SalidaAlmuerzo2") = Me.TxtSalida2.Text
        If Me.ChkUnicoHorario.Value = xtpChecked Then
         Me.AdoConsulta.Recordset("ExcluirSabado") = True
        Else
         Me.AdoConsulta.Recordset("ExcluirSabado") = False
        End If
        
        If Me.ChkAsignarHorario.Value = xtpChecked Then
          Me.AdoConsulta.Recordset("PersonalSinHorario") = True
        Else
          Me.AdoConsulta.Recordset("PersonalSinHorario") = False
        End If
        
        If Me.ChkHorasExtraLaborada.Value = xtpChecked Then
          Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "HorasTrab"
        Else
          Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "N"
        End If
        
        If Me.OptHorario.Value = True Then
           Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "Horario"
        End If
        
         If Me.OptHoras.Value = False Then
           Me.AdoConsulta.Recordset("CalcularHorasTrab") = 0
         Else
           Me.AdoConsulta.Recordset("CalcularHorasTrab") = 1
        End If
        
        If Me.TxtHorasTrab.Text <> "" Then
         Me.AdoConsulta.Recordset("HorasTrab") = Me.TxtHorasTrab.Text
        End If
        
        If Me.TxtHorasSabado.Text <> "" Then
         Me.AdoConsulta.Recordset("HorasTrabSab") = Me.TxtHorasSabado.Text
        End If
        
        If Me.TxtHorasDomingo.Text <> "" Then
         Me.AdoConsulta.Recordset("HorasTrabDom") = Me.TxtHorasDomingo.Text
        End If
        
        If Me.OptRestarAlmuerzo.Value = False Then
           Me.AdoConsulta.Recordset("RestarAlmuerzo") = 0
         Else
           Me.AdoConsulta.Recordset("RestarAlmuerzo") = 1
        End If
        
'        If Me.TxtHora.Text <> "  :  " Then
'         Me.AdoConsulta.Recordset("HoraSalida") = Me.TxtHora.Text
'        End If
        Me.AdoConsulta.Recordset.Update
   End If
   
   Me.AdoHorarios2.Refresh

MsgBox "Registro Grabado", vbExclamation, "Reportes Zeus"
End Sub

Private Sub CmdQuitarCredito_Click()
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

Private Sub Form_Load()
 MDIPrimero.Skin1.ApplySkin hWnd
 
  Me.CmdGuardar.BackColor = RGB(222, 227, 247)
  Me.CmdCerrar.BackColor = RGB(222, 227, 247)
  Me.CmdQuitarCredito.BackColor = RGB(222, 227, 247)
  
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
    

End Sub

Private Sub OptHoraEntrada_Click()
'  If Me.OptHoraEntrada.Value = True Then
'    Me.LblHora.Enabled = True
'    Me.LblHorasTrab.Enabled = True
'    Me.TxtHorasTrab.Enabled = True
'    Me.TxtHorasSabado.Enabled = True
'    Me.LblHoraDomingo.Enabled = True
'    Me.TxtHorasDomingo.Enabled = True
'  Else
'    Me.LblHora.Enabled = False
'    Me.LblHorasTrab.Enabled = False
'    Me.TxtHorasTrab.Enabled = False
'    Me.TxtHorasSabado.Enabled = False
'    Me.LblHoraDomingo.Enabled = False
'    Me.TxtHorasDomingo.Enabled = False
'  End If
End Sub

Private Sub OptHorario_Click()
 If Me.OptHorario.Value = True Then
    Me.LblHora.Enabled = False
    Me.LblHorasTrab.Enabled = False
    Me.TxtHorasTrab.Enabled = False
    Me.TxtHorasSabado.Enabled = False
    Me.LblHoraDomingo.Enabled = False
    Me.TxtHorasDomingo.Enabled = False
    Me.ChkHorasExtraLaborada.Enabled = False
    Me.ChkHorasExtraLaborada.Value = 0

 End If
End Sub

Private Sub OptHoras_Click()
  If Me.OptHoras.Value = True Then
    Me.LblHora.Enabled = True
    Me.LblHorasTrab.Enabled = True
    Me.TxtHorasTrab.Enabled = True
    Me.TxtHorasSabado.Enabled = True
    Me.LblHoraDomingo.Enabled = True
    Me.TxtHorasDomingo.Enabled = True
    Me.ChkHorasExtraLaborada.Enabled = True
  Else
    Me.LblHora.Enabled = False
    Me.LblHorasTrab.Enabled = False
    Me.TxtHorasTrab.Enabled = False
    Me.TxtHorasSabado.Enabled = False
    Me.LblHoraDomingo.Enabled = False
    Me.TxtHorasDomingo.Enabled = False
    Me.ChkHorasExtraLaborada.Enabled = False
  End If
End Sub

Private Sub OptHoraSalida_Click()
 If Me.OptHoraSalida.Value = True Then
    Me.LblHora.Enabled = False
    Me.LblHorasTrab.Enabled = False
    Me.TxtHorasTrab.Enabled = False
    Me.TxtHorasSabado.Enabled = False
    Me.LblHoraDomingo.Enabled = False
    Me.TxtHorasDomingo.Enabled = False
    Me.ChkHorasExtraLaborada.Enabled = False
    Me.ChkHorasExtraLaborada.Value = 0
 Else
    Me.LblHora.Enabled = True
    Me.LblHorasTrab.Enabled = True
    Me.TxtHorasTrab.Enabled = True
    Me.TxtHorasSabado.Enabled = True
    Me.LblHoraDomingo.Enabled = True
    Me.TxtHorasDomingo.Enabled = True
    Me.ChkHorasExtraLaborada.Enabled = True
 End If
End Sub

Private Sub TDBMatutino_Change()
  Dim Codigo As Double
  Codigo = Me.TDBMatutino.Columns(0).Text
  Me.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & Codigo & "))"
  Me.AdoConsulta.Refresh
  
  Sleep (1000)
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
        
        
        If Me.AdoConsulta.Recordset("ExcluirSabado") = True Then
          Me.ChkUnicoHorario.Value = xtpChecked
        Else
          Me.ChkUnicoHorario.Value = xtpUnchecked
        End If
        
        If Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "HorasTrab" Then
          Me.ChkHorasExtraLaborada.Value = xtpChecked
        Else
          Me.ChkHorasExtraLaborada.Value = xtpUnchecked
        End If
        
        If Me.AdoConsulta.Recordset("CalcularHorasTrab") = 0 Then
          Me.OptHoras.Value = False
          Me.OptHoraSalida.Value = True
        Else
          Me.OptHoras.Value = True
          Me.OptHoraSalida.Value = False
        End If
        
        If Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "Horario" Then
          Me.OptHoras.Value = False
          Me.OptHoraSalida.Value = False
          Me.OptHorario.Value = True
        End If
        
 
        If Not IsNull(Me.AdoConsulta.Recordset("HorasTrab")) Then
         Me.TxtHorasTrab.Text = Me.AdoConsulta.Recordset("HorasTrab")
        End If
        
        If Not IsNull(Me.AdoConsulta.Recordset("HorasTrabSab")) Then
         Me.TxtHorasSabado.Text = Me.AdoConsulta.Recordset("HorasTrabSab")
        End If
        
        If Not IsNull(Me.AdoConsulta.Recordset("HorasTrabDom")) Then
         Me.TxtHorasDomingo.Text = Me.AdoConsulta.Recordset("HorasTrabDom")
        End If
        
         If Me.AdoConsulta.Recordset("RestarAlmuerzo") = 0 Then
           Me.OptRestarAlmuerzo.Value = False
         Else
           Me.OptRestarAlmuerzo.Value = True
        End If
        
  End If
End Sub

Private Sub TDBMatutino_ItemChange()
  Dim Codigo As Double
  Codigo = Me.TDBMatutino.Columns(0).Text
  Me.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & Codigo & "))"
  Me.AdoConsulta.Refresh
'  Sleep (1000)
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
        
        If Me.AdoConsulta.Recordset("ExcluirSabado") = True Then
          Me.ChkUnicoHorario.Value = xtpChecked
        Else
          Me.ChkUnicoHorario.Value = xtpUnchecked
        End If
        
        If Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "HorasTrab" Then
          Me.ChkHorasExtraLaborada.Value = xtpChecked
        Else
          Me.ChkHorasExtraLaborada.Value = xtpUnchecked
        End If
        
        If Me.AdoConsulta.Recordset("CalcularHorasTrab") = 0 Then
          Me.OptHoras.Value = False
          Me.OptHoraSalida.Value = True
        Else
          Me.OptHoras.Value = True
          Me.OptHoraSalida.Value = False
        End If
        
        If Me.AdoConsulta.Recordset("TipoCalcularHorasTrab") = "Horario" Then
          Me.OptHorario.Value = True
        End If
 
        If Not IsNull(Me.AdoConsulta.Recordset("HorasTrab")) Then
         Me.TxtHorasTrab.Text = Me.AdoConsulta.Recordset("HorasTrab")
        End If
        
         
        If Not IsNull(Me.AdoConsulta.Recordset("HorasTrabSab")) Then
         Me.TxtHorasSabado.Text = Me.AdoConsulta.Recordset("HorasTrabSab")
        End If
        
        If Not IsNull(Me.AdoConsulta.Recordset("HorasTrabDom")) Then
         Me.TxtHorasDomingo.Text = Me.AdoConsulta.Recordset("HorasTrabDom")
        End If
        
'        If Not IsNull(Me.AdoConsulta.Recordset("HoraSalida")) Then
'         Me.TxtHora.Text = Format(Me.AdoConsulta.Recordset("HoraSalida"), "hh:mm")
'        End If
        
         If Me.AdoConsulta.Recordset("RestarAlmuerzo") = 0 Then
           Me.OptRestarAlmuerzo.Value = False
         Else
           Me.OptRestarAlmuerzo.Value = True
        End If
  End If
End Sub

