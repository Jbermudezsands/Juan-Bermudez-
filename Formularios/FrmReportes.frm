VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{60CCE6A8-5C61-4F30-8513-F57EED62E86A}#8.0#0"; "todl8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmReportesReloj 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   9870
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   4755
      Left            =   120
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   9720
      Begin VB.PictureBox picTV 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   960
         ScaleHeight     =   855
         ScaleWidth      =   7455
         TabIndex        =   4
         Top             =   960
         Width           =   7455
      End
      Begin VB.Label Lb9 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando........."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   20
         Top             =   2040
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label Lb1 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label Lb2 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb3 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb4 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando...."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   2040
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb5 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando....."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   15
         Top             =   2040
         Visible         =   0   'False
         Width           =   7095
      End
      Begin VB.Label Lb6 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando......"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   2040
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb7 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando......."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb8 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando........"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Label Lb0 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2040
         Width           =   6735
      End
      Begin VB.Label Lb10 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando.........."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   10
         Top             =   2040
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb11 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando..........."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb12 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando............"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb13 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb14 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando.............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2040
         UseMnemonic     =   0   'False
         Visible         =   0   'False
         Width           =   6735
      End
      Begin VB.Label Lb15 
         BackStyle       =   0  'Transparent
         Caption         =   "Procesando..............."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Visible         =   0   'False
         Width           =   5175
      End
      Begin VB.Image Img2 
         Height          =   480
         Left            =   1080
         Picture         =   "FrmReportes.frx":0000
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   495
      End
      Begin VB.Image img1 
         Height          =   480
         Left            =   2040
         Picture         =   "FrmReportes.frx":A487
         Stretch         =   -1  'True
         Top             =   1080
         Width           =   495
      End
   End
   Begin VB.CheckBox ChkLink 
      Caption         =   "Generar Reporte sin Link"
      Height          =   255
      Left            =   240
      TabIndex        =   65
      Top             =   5040
      Width           =   2535
   End
   Begin VB.CheckBox ChkRestarHorario 
      Caption         =   "Restar el tiempo del Horario"
      Height          =   315
      Left            =   4800
      TabIndex        =   63
      Top             =   7800
      Visible         =   0   'False
      Width           =   4695
   End
   Begin VB.CheckBox ChkJustificacion 
      Caption         =   "Empleados con Justificacion"
      Height          =   315
      Left            =   4320
      TabIndex        =   62
      Top             =   4920
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   4695
   End
   Begin XtremeSuiteControls.GroupBox FrameJustificacion 
      Height          =   735
      Left            =   10080
      TabIndex        =   55
      Top             =   3720
      Width           =   5295
      _Version        =   786432
      _ExtentX        =   9340
      _ExtentY        =   1296
      _StockProps     =   79
      Caption         =   "Tipo Justificacion"
      UseVisualStyle  =   -1  'True
      Begin VB.CommandButton Command5 
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
         Picture         =   "FrmReportes.frx":17E32
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command4 
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
         Picture         =   "FrmReportes.frx":17F80
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   240
         Width           =   375
      End
      Begin TrueOleDBList80.TDBCombo TDBJustificaIni 
         Bindings        =   "FrmReportes.frx":180CE
         Height          =   315
         Left            =   600
         TabIndex        =   58
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   7064
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
         ListField       =   "Classname"
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
         _PropDict       =   $"FrmReportes.frx":180ED
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
      Begin TrueOleDBList80.TDBCombo TDBJustificaFin 
         Bindings        =   "FrmReportes.frx":18197
         Height          =   315
         Left            =   3240
         TabIndex        =   59
         Top             =   240
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _LayoutType     =   0
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         _DropdownWidth  =   7064
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
         ListField       =   "Classname"
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
         _PropDict       =   $"FrmReportes.frx":181B6
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
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin"
         Height          =   255
         Left            =   2880
         TabIndex        =   60
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CheckBox ChkOrdenado 
      Caption         =   "Ordenado por Codigo"
      Height          =   315
      Left            =   4320
      TabIndex        =   53
      Top             =   4560
      Visible         =   0   'False
      Width           =   4695
   End
   Begin SmartButtonProject.SmartButton CmdVerReporte 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":18260
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
   Begin SmartButtonProject.SmartButton CmdVerReporte2 
      Height          =   855
      Left            =   120
      TabIndex        =   47
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":19DB4
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
   Begin SmartButtonProject.SmartButton CmdVerReporte3 
      Height          =   855
      Left            =   120
      TabIndex        =   52
      Top             =   5880
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":1B908
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
   Begin MSAdodcLib.Adodc AdoDepartamento 
      Height          =   375
      Left            =   4440
      Top             =   10080
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   11280
      Top             =   8400
      _ExtentX        =   847
      _ExtentY        =   847
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
      TabIndex        =   22
      Top             =   0
      Width           =   9975
      Begin VB.Image Image2 
         Height          =   960
         Left            =   360
         Picture         =   "FrmReportes.frx":1D45C
         Stretch         =   -1  'True
         Top             =   80
         Width           =   1080
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   0
         X2              =   9960
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Image Image1 
         Height          =   645
         Left            =   480
         Top             =   120
         Width           =   645
      End
      Begin VB.Label LblTitulo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reportes Generales"
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
         Left            =   2760
         TabIndex        =   23
         Top             =   360
         Width           =   4320
      End
   End
   Begin VB.Timer Timer1 
      Left            =   10560
      Top             =   8280
   End
   Begin SmartButtonProject.SmartButton CmdSalir 
      Height          =   855
      Left            =   8520
      TabIndex        =   2
      Top             =   5880
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Salir"
      Picture         =   "FrmReportes.frx":1DB88
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
   Begin MSComDlg.CommonDialog CDRuta 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin XtremeSuiteControls.ProgressBar osProgress2 
      Height          =   255
      Left            =   4200
      TabIndex        =   25
      Top             =   6480
      Visible         =   0   'False
      Width           =   4215
      _Version        =   786432
      _ExtentX        =   7435
      _ExtentY        =   450
      _StockProps     =   93
      Appearance      =   6
   End
   Begin XtremeSuiteControls.ProgressBar osProgress1 
      Height          =   375
      Left            =   1560
      TabIndex        =   24
      Top             =   6000
      Visible         =   0   'False
      Width           =   6855
      _Version        =   786432
      _ExtentX        =   12091
      _ExtentY        =   661
      _StockProps     =   93
      Scrolling       =   1
      Appearance      =   6
   End
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   4575
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Width           =   9615
      _Version        =   786432
      _ExtentX        =   16960
      _ExtentY        =   8070
      _StockProps     =   68
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.DisableLunaColors=   0   'False
      PaintManager.OneNoteColors=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      ItemCount       =   1
      Item(0).Caption =   "Reportes"
      Item(0).ControlCount=   8
      Item(0).Control(0)=   "CmbReportes"
      Item(0).Control(1)=   "Label22"
      Item(0).Control(2)=   "LblMonto"
      Item(0).Control(3)=   "FrameDpto"
      Item(0).Control(4)=   "FrameFecha"
      Item(0).Control(5)=   "FrameEmpleado"
      Item(0).Control(6)=   "ChkAcumulado"
      Item(0).Control(7)=   "ChkTodosDptos"
      Begin VB.CheckBox ChkTodosDptos 
         Caption         =   "Incluir Todos los Departamentos"
         Height          =   315
         Left            =   4200
         TabIndex        =   51
         Top             =   4080
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CheckBox ChkAcumulado 
         Caption         =   "Calcular Acumulado Rango de Fechas"
         Height          =   315
         Left            =   4080
         TabIndex        =   50
         Top             =   1440
         Visible         =   0   'False
         Width           =   4695
      End
      Begin XtremeSuiteControls.GroupBox FrameDpto 
         Height          =   735
         Left            =   4080
         TabIndex        =   30
         Top             =   600
         Width           =   5295
         _Version        =   786432
         _ExtentX        =   9340
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Departamento"
         UseVisualStyle  =   -1  'True
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
            Picture         =   "FrmReportes.frx":1EE9A
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   240
            Width           =   375
         End
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
            Picture         =   "FrmReportes.frx":1EFE8
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   240
            Width           =   375
         End
         Begin TrueOleDBList80.TDBCombo DBDptoIni 
            Bindings        =   "FrmReportes.frx":1F136
            Height          =   315
            Left            =   600
            TabIndex        =   33
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   556
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   7064
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
            DropdownPosition=   1
            Locked          =   0   'False
            ScrollTrack     =   0   'False
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            AddItemSeparator=   ";"
            _PropDict       =   $"FrmReportes.frx":1F154
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
            Bindings        =   "FrmReportes.frx":1F1FE
            Height          =   315
            Left            =   3240
            TabIndex        =   34
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _LayoutType     =   0
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            _DropdownWidth  =   7064
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
            DropdownPosition=   1
            Locked          =   0   'False
            ScrollTrack     =   0   'False
            RowDividerColor =   14215660
            RowSubDividerColor=   14215660
            AddItemSeparator=   ";"
            _PropDict       =   $"FrmReportes.frx":1F21C
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
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Fin"
            Height          =   255
            Left            =   2880
            TabIndex        =   32
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.ListBox CmbReportes 
         Height          =   3180
         ItemData        =   "FrmReportes.frx":1F2C6
         Left            =   120
         List            =   "FrmReportes.frx":1F2C8
         TabIndex        =   29
         Top             =   600
         Width           =   3735
      End
      Begin XtremeSuiteControls.GroupBox FrameFecha 
         Height          =   735
         Left            =   4080
         TabIndex        =   35
         Top             =   600
         Visible         =   0   'False
         Width           =   5295
         _Version        =   786432
         _ExtentX        =   9340
         _ExtentY        =   1296
         _StockProps     =   79
         UseVisualStyle  =   -1  'True
         Begin MSComCtl2.DTPicker DtpFechaINI 
            Height          =   300
            Left            =   600
            TabIndex        =   38
            Top             =   240
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            Format          =   77856769
            CurrentDate     =   40789
         End
         Begin MSComCtl2.DTPicker DTFechaFin 
            Height          =   300
            Left            =   3360
            TabIndex        =   39
            Top             =   240
            Width           =   1815
            _ExtentX        =   3201
            _ExtentY        =   529
            _Version        =   393216
            Format          =   77856769
            CurrentDate     =   40789
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Fin"
            Height          =   255
            Left            =   2880
            TabIndex        =   37
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   240
            Width           =   615
         End
      End
      Begin XtremeSuiteControls.GroupBox FrameEmpleado 
         Height          =   735
         Left            =   4080
         TabIndex        =   40
         Top             =   1800
         Visible         =   0   'False
         Width           =   5295
         _Version        =   786432
         _ExtentX        =   9340
         _ExtentY        =   1296
         _StockProps     =   79
         Caption         =   "Registros de Empleados"
         UseVisualStyle  =   -1  'True
         Begin VB.CommandButton Command3 
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
            Picture         =   "FrmReportes.frx":1F2CA
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton Command2 
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
            Picture         =   "FrmReportes.frx":1F418
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   240
            Width           =   375
         End
         Begin TrueOleDBList80.TDBCombo TDBCombo1 
            Bindings        =   "FrmReportes.frx":1F566
            Height          =   315
            Left            =   600
            TabIndex        =   41
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
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
            _PropDict       =   $"FrmReportes.frx":1F582
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
         Begin TrueOleDBList80.TDBCombo DBEmpleado2 
            Bindings        =   "FrmReportes.frx":1F62C
            Height          =   315
            Left            =   3240
            TabIndex        =   42
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
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
            _PropDict       =   $"FrmReportes.frx":1F648
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
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Inicio"
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Fin"
            Height          =   255
            Left            =   2880
            TabIndex        =   43
            Top             =   240
            Width           =   615
         End
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
         TabIndex        =   28
         Top             =   5040
         Visible         =   0   'False
         Width           =   3135
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
         TabIndex        =   27
         Top             =   5040
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   600
      Top             =   9360
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
      Left            =   600
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
      Left            =   600
      Top             =   10440
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
      Left            =   4440
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
      Left            =   4320
      Top             =   10560
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
      Left            =   600
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
   Begin MSAdodcLib.Adodc AdoHorarioAlmuerzo 
      Height          =   375
      Left            =   4560
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
      Left            =   4920
      Top             =   8160
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
   Begin SmartButtonProject.SmartButton CmdVerReporte4 
      Height          =   855
      Left            =   1920
      TabIndex        =   54
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":1F6F2
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
   Begin MSAdodcLib.Adodc AdoJustificacion 
      Height          =   375
      Left            =   360
      Top             =   8160
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
   Begin SmartButtonProject.SmartButton CmdVerReporte5 
      Height          =   855
      Left            =   120
      TabIndex        =   64
      Top             =   7920
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1508
      ForeColor       =   8388608
      Caption         =   "Ver Reporte"
      Picture         =   "FrmReportes.frx":21246
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
   Begin MSAdodcLib.Adodc AdoConsultaReporte 
      Height          =   375
      Left            =   2760
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
   Begin VB.Label lblProgreso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   1200
      TabIndex        =   21
      Top             =   8520
      Width           =   45
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reportes"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   1
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "FrmReportesReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim Tape As New clsTape
Dim HayUtiBruta As Boolean



Private Sub ChkExportar_Click()
' If Me.ChkExportar.Value = 1 Then
'   Me.CommonDialog1.ShowSave
'  RutaArchivo = ""
'  RutaArchivo = Me.CommonDialog1.FileName + ".xls"
'
' End If
End Sub




Private Sub ChkTodosDptos_Click()
 If Me.ChkTodosDptos.Value = 1 Then
   Me.DBDptoIni.Text = ""
   Me.DBDptoFin.Text = ""
   Me.TDBCombo1.Text = ""
   Me.DBEmpleado2.Text = ""
 
 End If

End Sub

Private Sub CmbReportes_Click()
Dim FechaFin As Double

Me.FrameDpto.Visible = False
Me.FrameFecha.Visible = False
Me.FrameEmpleado.Visible = False
Me.FrameEmpleado.Top = 1440
Me.DTFechaFin.Enabled = True
Me.CmdVerReporte.Visible = True
Me.CmdVerReporte2.Visible = False
Me.CmdVerReporte3.Visible = False
Me.CmdVerReporte4.Visible = False
Me.CmdVerReporte5.Visible = False
Me.ChkAcumulado.Visible = False
Me.ChkTodosDptos.Visible = False
Me.CmdVerReporte3.Visible = False
Me.ChkOrdenado.Visible = False
Me.ChkRestarHorario.Visible = False

     Me.CmdVerReporte.Left = 120
     Me.CmdVerReporte.Top = 6000
'     Me.CmdVerReporte1.Left = 120
'     Me.CmdVerReporte1.Top = 6000
     Me.CmdVerReporte2.Left = 120
     Me.CmdVerReporte2.Top = 6000
     Me.CmdVerReporte3.Left = 120
     Me.CmdVerReporte3.Top = 6000
     Me.CmdVerReporte4.Left = 120
     Me.CmdVerReporte4.Top = 6000
     Me.CmdVerReporte5.Left = 120
     Me.CmdVerReporte5.Top = 6000

Select Case Me.CmbReportes.Text
 Case "REPORTE ASISTENCIA CATORCE DIAS"
    Me.DTFechaFin.Enabled = False

   Me.FrameFecha.Visible = True
   FechaFin = DateAdd("D", 13, Me.DTPFechaIni.Value)
   Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.CmdVerReporte5.Visible = True
     Me.CmdVerReporte5.Left = 120
     Me.CmdVerReporte5.Top = 6000
     Me.ChkJustificacion.Visible = False
     Me.ChkTodosDptos.Visible = True
     
 Case "REPORTE ASISTENCIA QUINCE DIAS"
    Me.DTFechaFin.Enabled = False

   Me.FrameFecha.Visible = True
   FechaFin = DateAdd("D", 14, Me.DTPFechaIni.Value)
   Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.CmdVerReporte5.Visible = True
     Me.CmdVerReporte5.Left = 120
     Me.CmdVerReporte5.Top = 6000
         Me.ChkJustificacion.Visible = False
     Me.ChkTodosDptos.Visible = True
     
 Case "REPORTE MARCAS INCOMPLETAS"
   Me.FrameFecha.Visible = True
   FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
   Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.CmdVerReporte4.Visible = True
     Me.CmdVerReporte4.Left = 120
     Me.CmdVerReporte4.Top = 6000
         Me.ChkJustificacion.Visible = False
     Me.ChkTodosDptos.Visible = True



Case "REPORTE AUSENCIA X DIA"
     Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.CmdVerReporte4.Visible = True
     Me.CmdVerReporte4.Left = 120
     Me.CmdVerReporte4.Top = 6000
     Me.ChkJustificacion.Visible = True
     Me.ChkTodosDptos.Visible = True

Case "REPORTE TARJETA DE ASISTENCIA X DIA"
     Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 1, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.CmdVerReporte3.Visible = False
     Me.CmdVerReporte4.Visible = True
     Me.CmdVerReporte4.Left = 120
     Me.CmdVerReporte4.Top = 6000
     
     Me.ChkTodosDptos.Visible = True
Case "REPORTE TARJETA DE ASISTENCIA X DIA"
     Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 1, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.CmdVerReporte3.Visible = True
     
     Me.ChkTodosDptos.Visible = True


Case "REPORTE DE ENTRADAS Y SALIDAS"
     Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.CmdVerReporte3.Visible = True
     
     Me.ChkTodosDptos.Visible = True

 Case "REPORTE DE JUSTIFICACION"
      Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2200
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1400
     Me.CmdVerReporte3.Visible = False
     Me.CmdVerReporte2.Visible = True
     Me.CmdVerReporte.Visible = False
     Me.ChkTodosDptos.Visible = True
     Me.ChkTodosDptos.Top = 3840
     Me.FrameJustificacion.Top = 4200
     Me.FrameJustificacion.Left = 4150
     Me.FrameJustificacion.Visible = True

 Case "LISTADO EMPLEADOS"
   Me.FrameDpto.Visible = True
   FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
   Me.DTFechaFin.Value = FechaFin

   
 Case "REPORTE ASISTENCIA X DIA"
   Me.FrameFecha.Visible = True
   FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
   Me.DTFechaFin.Value = FechaFin
   Me.FrameDpto.Top = 1560
   Me.FrameDpto.Visible = True
   Me.ChkAcumulado.Top = 1440
   Me.ChkAcumulado.Visible = True
   Me.FrameDpto.Top = 1750
   Me.ChkTodosDptos.Visible = True
 Case "REPORTE LLEGADAS TARDE"
   Me.FrameFecha.Visible = True
   FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
   Me.DTFechaFin.Value = FechaFin
   Me.FrameDpto.Top = 1560
   Me.FrameDpto.Visible = True
   
 Case "REPORTE SALIDA ANTICIPADA"
   Me.FrameFecha.Visible = True
   FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
   Me.DTFechaFin.Value = FechaFin
   Me.FrameDpto.Top = 1560
   Me.FrameDpto.Visible = True
   
   
 Case "REPORTE ASISTENCIA SIETE DIAS"
    Me.FrameFecha.Visible = True
    Me.DTFechaFin.Enabled = False
    FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
    Me.DTFechaFin.Value = FechaFin
    Me.FrameDpto.Visible = True
    Me.FrameDpto.Top = 1400
    Me.ChkOrdenado.Visible = True
 Case "REPORTE HORAS LAB EXTRA SIETE DIAS"
     Me.FrameFecha.Visible = True
     Me.DTFechaFin.Enabled = False
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1400
     Me.CmdVerReporte3.Visible = True
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
     Me.ChkTodosDptos.Visible = True
     Me.ChkRestarHorario.Visible = True
     Me.ChkRestarHorario.Top = 3840
     
     
 Case "REPORTE HORAS LABORADAS SIETE DIAS"
     Me.FrameFecha.Visible = True
     Me.DTFechaFin.Enabled = False
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1400
 Case "REPORTE HORAS EXTRA SIETE DIAS"
     Me.FrameFecha.Visible = True
     Me.DTFechaFin.Enabled = False
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1400
     Me.CmdVerReporte3.Visible = False
     Me.CmdVerReporte2.Visible = True
     Me.CmdVerReporte.Visible = False
 Case "REPORTE LLEGADAS TARDE SIETE DIAS"
     Me.FrameFecha.Visible = True
     Me.DTFechaFin.Enabled = False
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1400
     Me.CmdVerReporte3.Visible = True
     Me.CmdVerReporte2.Visible = False
     Me.CmdVerReporte.Visible = False
 Case "REPORTE DETALLE ASISTENCIA"
     Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.CmdVerReporte2.Visible = True
     Me.CmdVerReporte.Visible = False
 Case "REPORTE ASISTENCIA Y AUSENCIA X DIA"
     Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
     Me.FrameEmpleado.Visible = True
     Me.FrameEmpleado.Top = 2400
     Me.FrameDpto.Visible = True
     Me.FrameDpto.Top = 1440
     Me.CmdVerReporte2.Visible = True
     Me.CmdVerReporte.Visible = False
     Me.ChkTodosDptos.Visible = True
End Select
End Sub

Private Sub CmdBuscaCuenta_Click()
'  Quien = "DptoIni"
'  MDIPrimero.MousePointer = 11
'  FrmDepartamentoReportes.Show 1
'  MDIPrimero.MousePointer = 0
  Quien = "Departamento"
  FrmConsultaReloj.Show 1
  Me.DBDptoIni.Text = NombreDepartamento
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub



Private Sub CmdVerReporte_Click()
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
Dim EntradaA As String, SalidaA As String, HoraAlmuerzo As String, DifHoras As Double, EntradaAlmuerzo As String, SalidaAlmuerzo As String, EntradaAlmuerzo1 As String, EntradaAlmuerzo2 As String, SalidaAlmuerzo1 As String, SalidaAlmuerzo2 As String, ExcluirSabado As Boolean
Dim SQlEntradaAlmuerzo As String, SqlSalidaAlmuerzo As String
TieneJornadas = False
Me.osProgress2.Visible = False
SinHorario = True

 If Not IsNull(Me.AdoDatosEmpresa.Recordset("MinutosExtra")) Then
  MinutosExtra = Me.AdoDatosEmpresa.Recordset("MinutosExtra")
 Else
  MinutosExtra = 0
 End If
 
 CantHorarios = 0

      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj

Me.Frame6.Visible = True
Me.Frame6.Top = 1080


Select Case Me.CmbReportes.Text
 Case "REPORTE HORAS LABORADAS SIETE DIAS"
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
'      SQL = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      Else
       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
             "WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"
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
        
        
                If CodEmpleado = "1057" Then
                  CodEmpleado = "1057"
                End If
        
        b = 1
        
          Me.osProgress2.Visible = True
          Me.osProgress2.Max = 6
          Me.osProgress2.Min = 0
          Me.osProgress2.Value = 0
          
          TotalHorasTrabajadas = 0
          TotalTrabajadas = "00:00"
        
          For j = 0 To 6

                 If j = 0 Then
                    '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                    sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                          "WHERE ((UserShift.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Me.DTFechaFin.Value, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Me.DTPFechaIni.Value, "YYYY-MM-DD") & "')"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                      Ciclo = Me.AdoConsulta.Recordset("Cycles")
                      Date1 = CDate(FechaInicioH)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                    Else
                      Date1 = CDate(Me.DTPFechaIni.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                   
                    End If
                 Else
                        If FechaInicioH <> "" Then
                         Date1 = CDate(FechaInicioH)
                        Else
                         Date1 = CDate(Me.DTPFechaIni)
                        End If
'                        Date1 = CDate(Me.DtpFechaINI)
                        Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                        DiaInicio = DiaHorario(Date1, Date2, Ciclo)

                End If
                
                Me.Caption = "Procesando " & Date2 & " Empleado: " & i & " de " & Me.osProgress1.Max
                DoEvents
                
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
                 Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                    If Not Me.AdoHorarios.Recordset.EOF Then
                    
                    
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Date2, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Date2, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                        
                        LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                          If LongitudMinutosIn < 1200 Then  'MENOR A 1400MIN 12 HORAS
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                           
                           SinHorario = True
                           SqlIN(0) = sql
                           SqlOut(0) = SQlSalida
                           CantHorarios = 1
                      Else
                      
                           TieneJornadas = False
                           SinHorario = False
                           CantHorarios = 0
                           Me.AdoHorarios.Refresh
                        
                           Do While Not Me.AdoHorarios.Recordset.EOF
                           
                               BInTime = Me.AdoHorarios.Recordset("BIntime")
                               EInTime = Me.AdoHorarios.Recordset("EIntime")
                               InTime = Me.AdoHorarios.Recordset("Intime")
                               LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                               
'                               Me.AdoHorarios.Recordset.MoveLast
                               
                               BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                               EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                               OutTime = Me.AdoHorarios.Recordset("OutTime")
                               If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                               Else
                                TardePermintido = 0
                               End If
                               
                               
                               FechaIn = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               FechaOut = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               
                               FechaHInicio = "#" & FechaIn & " " & BInTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EInTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BInTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                               
                               sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                
                               FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                          
                           
    
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & InTime))
'                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                               
                                ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                               LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & EInTime))
                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                              FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                               FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                               
                               SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                
                                
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaInicio, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                    End If
                                
                                
                                '********************************************************************************************
                                '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                '********************************************************************************************
                                
                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                CodigoH = Me.AdoHorarios.Recordset("Schid")
                            
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
                                    
                       
                       
                       
                       
                       SqlIN(CantHorarios) = sql
                       SqlOut(CantHorarios) = SQlSalida
                       CantHorarios = CantHorarios + 1
                       Me.AdoHorarios.Recordset.MoveNext
                     Loop
                     End If
                    
                 Else '//////SI NO TIENE HORARIO SOLO AGREGO LOS REGISTROS DE ENTRADA ///////////
                    
                       
                       FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & "#"
                       FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59:59#"
                       
                       BInTime = ""
                       EInTime = ""
                       InTime = ""
                       
        '               Me.AdoHorarios.Recordset.MoveLast
                       
                       BOutTime = ""
                       EOutTime = ""
                       OutTime = ""
                       
                       
                      '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                      Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                      Me.AdoBuscaReporte.Refresh
                      If Not Me.AdoBuscaReporte.Recordset.EOF Then
                          CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                          HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                          RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                          RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                          JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                          
                         
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                      End If
                        SinHorario = True
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                    End If
                
               For L = 0 To CantHorarios - 1
                        sql = SqlIN(L)
                        SQlSalida = SqlOut(L)
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
                            '///////////////CON ESTA CONSULTA BUSCO LA HORA DE SALIDA///////////////////////////////////
                            '*********************************************************************************************
                                Salida = "00:00"
                                If TieneJornadas = True Then
                                   
                                     '///////////////////////////////CON ESTAS FECHAS BUSCO LA HORA DE SALIDA DE LA JORNADA ///////////////////
                                     
                                     
                                     HoraSalida = CDate(Entrada) + CDate(CInt(HorasLaborales) & ":00:00")
                                     FechaHInicio = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") - CDate(RangoHora1 & ":00")), "mm/dd/yyyy hh:mm:ss")
                                     FechaHFinal = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") + CDate(RangoHora2 & ":00")), "mm/dd/yyyy hh:mm:ss")
'                                     HoraSalida = Format(Date2, "mm/dd/yyyy") & " 23:59:59"
                                     HoraSalida = FechaHFinal
                                     HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                     If JornadaIntercalada = False Then
'                                        If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                           FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                        End If
'                                     End If
                               
                                    FechaHInicio = "#" & FechaHInicio & "#"
                                    FechaHFinal = "#" & FechaHFinal & "#"
                                    
                                    SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                                      If Entrada = Salida Then
                                        Entrada = "00:00"
                                        Salida = "00:00"
                                      End If
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
                            End If
                            
                            If CodEmpleado = 1057 Then
                              CodEmpleado = 1057
                            End If
                      
                            
                            '*********************************************************************************************
                            '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                            '*********************************************************************************************
                            RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                            
                            If Entrada = "00:00" Then
                              Salida = "00:00"
                            End If
                            
                                HorasTrabajadas = 0
                                HoraLaboradas = "00:00"
                                If Salida <> "00:00" Then
                                 If Entrada <> "00:00" Then
            '                      HorasTrabajadas = (DateDiff("h", Entrada, Salida))
                                   HoraLaboradas = ConvertirSegundos((DateDiff("s", Entrada, Salida)), DiaInicio)
                                   HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60) - RestarAlmuerzo  '/////RESTO UNA HORA DE ALMUERZO //////
                                   HoraSalida = Format(Salida, "hh:mm:ss")
                                   TotalTrabajadas = HoraLaboradas + TotalTrabajadas
            
                                 
                                 Else
                                    HorasTrabajadas = 0
                                    HoraLaboradas = "00:00"
                                 End If
                                End If
            
                            
            '                If Salida <> "00:00" Then
            '                  HoraLaboradas = (DateDiff("n", Entrada, Salida) / 60)
            '                  HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60)
            '                Else
            '                   HoraLaboradas = 0
            '                   HorasTrabajadas = 0
            '                End If
                            
                            
                              
                              
                              
            '                HoraSalida = Format(Salida, "hh:mm:ss")
            '                If OutTime <> "?" Then
            '                 HoraSalidaHorario = OutTime
            '                 HorasExtras = (DateDiff("n", HoraSalidaHorario, HoraSalida) / 60)
            '                Else
            '                 HorasExtras = 0
            '                End If
            '
            '                If HorasTrabajadas < 0 Then
            '                  HorasTrabajadas = 0
            '                End If
            '
            '                If HorasExtras < 0 Then
            '                  HorasExtras = 0
            '                End If

                             '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                             '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                             '/////////////////////////////////////////////////////////////////////////////////////
                             Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                             Me.AdoConsulta.Refresh
                             If Not Me.AdoConsulta.Recordset.EOF Then
                            
                                        Select Case j
                                        
                                            Case 0
                                                Me.AdoReportes.Recordset.AddNew
                                                 Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                                 Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                                 Me.AdoReportes.Recordset("Campo3") = departamento
                                                 Me.AdoReportes.Recordset("Campo6") = HoraLaboradas
                                                 Me.AdoReportes.Recordset.Update
                                                Me.AdoReportes.Refresh
                                             Case 1
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo7") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                                 
                                             Case 2
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo8") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 3
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo9") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 4
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo10") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 5
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo11") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 6
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo12") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset("Campo5") = Format(TotalTrabajadas, "hh:mm")
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                        End Select
                                End If
                Next
                Me.osProgress2.Value = j + 1
        
           Next
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
        Me.AdoBuscaReporte.Refresh
      Loop
      
         
      
         sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.Campo6 AS Dia1, Reportes.Campo7 AS Dia2, Reportes.Campo8 AS Dia3, Reportes.Campo9 AS Dia4, Reportes.Campo10 AS Dia5, Reportes.Campo11 AS Dia6, Reportes.Campo12 AS Dia7, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7, Reportes.CampoNum1 AS TotalHoras,Reportes.Campo5 From Reportes ORDER BY Reportes.Campo3,Reportes.Campo1,Reportes.CampoFecha1"


         Set rpt = New ArepHorasLaboradas
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
 
 Case "REPORTE ASISTENCIA SIETE DIAS"
     
      
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      
      If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
        'sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) "
         sql = "SELECT DISTINCT Checkinout.Userid, Userinfo.Name FROM Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) "

      Else
'       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
'             "WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"

       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName, Userinfo.Name FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
             "WHERE (((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) "

      End If
      
      
      If Me.ChkOrdenado.Value = 1 Then
        sql = sql & " ORDER BY Checkinout.Userid"
      Else
        sql = sql & " ORDER BY Userinfo.Name"
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
      
     

      FechaInicial = Me.DTPFechaIni.Value
      
      Do While Not Me.AdoEmpleados.Recordset.EOF
        DoEvents
        
        CodEmpleado = Me.AdoEmpleados.Recordset("Userid")
        
                If CodEmpleado = "168" Then
                  CodEmpleado = "168"
                End If
        b = 1
        
          Me.osProgress2.Visible = True
          Me.osProgress2.Max = 6
          Me.osProgress2.Min = 0
          Me.osProgress2.Value = 0
        
          For j = 0 To 6

                 If j = 0 Then
                    '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                    sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                          "WHERE ((UserShift.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Me.DTFechaFin.Value, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Me.DTPFechaIni.Value, "YYYY-MM-DD") & "')"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                      Ciclo = Me.AdoConsulta.Recordset("Cycles")
                      Date1 = CDate(FechaInicioH)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                    Else
                    
                      Date1 = CDate(Me.DTPFechaIni.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                    End If
                 Else
                      If FechaInicioH = "" Then
                        Date1 = CDate(Me.DTPFechaIni.Value)
                      Else
                        Date1 = CDate(FechaInicioH)
                      End If
'                      Date1 = CDate(Me.DtpFechaINI.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                End If
                
                Me.Caption = "Procesando " & Date2 & " Empleado: " & i & " de " & Me.osProgress1.Max
                DoEvents
                
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
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                  If Not Me.AdoHorarios.Recordset.EOF Then
                    
                    
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(Date2, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Date2, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                        
                        LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                          If LongitudMinutosIn < 720 Then  'MENOR A 1400MIN 12 HORAS
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                        SinHorario = True
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                      Else
                       SinHorario = False
                       TieneJornadas = False
                       CantHorarios = 0
                       Me.AdoHorarios.Refresh
                        
                       Do While Not Me.AdoHorarios.Recordset.EOF
                       
                               BInTime = Me.AdoHorarios.Recordset("BIntime")
                               EInTime = Me.AdoHorarios.Recordset("EIntime")
                               InTime = Me.AdoHorarios.Recordset("Intime")
                               LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                               
'                               Me.AdoHorarios.Recordset.MoveLast
                               
                               BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                               EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                               OutTime = Me.AdoHorarios.Recordset("OutTime")
                               If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                               Else
                                TardePermintido = 0
                               End If
                               
                               
                               FechaIn = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               FechaOut = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               
                               FechaHInicio = "#" & FechaIn & " " & BInTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EInTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BInTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                               
                               sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                               
                               '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                               '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                               '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaIn & " " & InTime))
'                               FechaHInicio = "#" & Mid(CDate(HorasIn), 1, 10) & " " & BOutTime & "#" 'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & Mid(CDate(HorasIn), 1, 10) & " " & EOutTime & "#"
                               
        '                       FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        ''                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
        '                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
        '                       MinutosTarde = MinutosSalida & ":00" & ":00"
        '                       FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
        '                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                               
                               FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                          
                         
    
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & InTime))
'                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                               
                               ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                               LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & EInTime))
                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                              FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                               FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                               
                               SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                           "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                
                            
                                '********************************************************************************************
                                '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                '********************************************************************************************
                                
                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
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
                                
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaInicio, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                    End If
                                
                         
                       SqlIN(CantHorarios) = sql
                       SqlOut(CantHorarios) = SQlSalida
                       CantHorarios = CantHorarios + 1
                       Me.AdoHorarios.Recordset.MoveNext
                     Loop
                    End If
                      
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
                       
                         '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                         Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                         Me.AdoBuscaReporte.Refresh
                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                             CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                             HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                             RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                             RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                             JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                             
                            
                             
                             sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                           
                             SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                             
                             TieneJornadas = True
                        
                         Else
                         
                             TieneJornadas = False
                             sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)= " & Ent & "))"  'I
                           
                             SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)= " & Sal & "))"  'O
                         End If
                         
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                  End If
                

                
                '*********************************************************************************************
                '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                '*********************************************************************************************
               
                For L = 0 To CantHorarios - 1
                        sql = SqlIN(L)
                        SQlSalida = SqlOut(L)
                            
                        If CodEmpleado = "2078" Then
                          CodEmpleado = "2078"
                        End If
                            
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
                            
                           If Entrada <> "00:00" Then
                              If ConfCalcularHorasTrab = True Then
                                  If TipoHorasTrabajada = "HorasTrab" Then
                                    If InTime <> "?" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
                                    End If
                                  End If
                              End If
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
                             HoraSalida = Format(Date2, "mm/dd/yyyy") & " 23:59:59"
                             HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                             If JornadaIntercalada = False Then
                                If CDate(FechaHFinal) > CDate(HoraSalida) Then
                                   FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                                End If
                             End If
                       
                            FechaHInicio = "#" & FechaHInicio & "#"
                            FechaHFinal = "#" & FechaHFinal & "#"
                            
                            SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
    '                HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60)
    '                HoraSalida = Format(Salida, "hh:mm:ss")
    '                If OutTime <> "?" Then
    '                 HoraSalidaHorario = OutTime
    '                 HorasExtras = (DateDiff("n", HoraSalidaHorario, HoraSalida) / 60)
    '                Else
    '                 HorasExtras = 0
    '                End If
    '
    '                If HorasTrabajadas < 0 Then
    '                  HorasTrabajadas = 0
    '                End If
    '
    '                If HorasExtras < 0 Then
    '                  HorasExtras = 0
    '                End If
    
                        If CodEmpleado = "168" Then
                         CodEmpleado = "168"
                        End If
                     
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                        '/////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                    
                            Select Case j
                            
                                Case 0
                                '-------------------------BUSCO SI EL EMPLEADO YA TIENE REGISTROS --------------------------------
                                    Me.AdoConsultaReporte.RecordSource = "SELECT Reportes.* From Reportes WHERE (((Reportes.Campo1)='" & CodEmpleado & "'))"
                                    Me.AdoConsultaReporte.Refresh
                                    If Me.AdoConsultaReporte.Recordset.EOF Then
                                        Me.AdoReportes.Recordset.AddNew
                                         Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                         Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                         Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                         Me.AdoReportes.Recordset("Campo3") = departamento
                                         Me.AdoReportes.Recordset("CampoFecha1") = Entrada
                                         Me.AdoReportes.Recordset("CampoFecha2") = Salida
                                        Me.AdoReportes.Recordset.Update
                                        Me.AdoReportes.Refresh
                                    End If
                                 Case 1
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha3") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha4") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                     
                                 Case 2
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha5") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha6") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 3
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha7") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha8") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 4
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha9") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha10") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 5
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha11") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha12") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 6
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha13") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha14") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                            End Select
                      End If
                Next
                Me.osProgress2.Value = j + 1
                DoEvents
           Next
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max & "Empleado " & CodEmpleado
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         If Me.ChkOrdenado.Value = 1 Then
          sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada1, Reportes.CampoFecha2 AS Salida1, Reportes.CampoFecha3 AS Entrada2, Reportes.CampoFecha4 AS Salida2, Reportes.CampoFecha5 AS Entrada3, Reportes.CampoFecha6 AS Salida3, Reportes.CampoFecha7 AS Entrada4, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7 From Reportes ORDER BY Reportes.Campo3,Reportes.CampoNum1,Reportes.CampoFecha1"
         Else
          sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada1, Reportes.CampoFecha2 AS Salida1, Reportes.CampoFecha3 AS Entrada2, Reportes.CampoFecha4 AS Salida2, Reportes.CampoFecha5 AS Entrada3, Reportes.CampoFecha6 AS Salida3, Reportes.CampoFecha7 AS Entrada4, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7 From Reportes ORDER BY Reportes.Campo3,Reportes.Campo2, Reportes.CampoNum1,Reportes.CampoFecha1"
         End If
         
         Set rpt = New ArepAsistenciaSiete
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
 Case "LISTADO EMPLEADOS"
      
      If Me.DBDptoIni.Text = "" Or Me.DBDptoFin.Text = "" Then
     
        sql = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName, Userinfo.Nation, Userinfo.Polity FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid ORDER BY Userinfo.Cardnum"
        Set rpt = New ArpListadoEmpleados
      Else
       CodDptoIni = Me.DBDptoIni.Columns(0).Text
       CodDptoFin = Me.DBDptoFin.Columns(0).Text
       sql = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName, Userinfo.Nation, Userinfo.Polity FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Dept.Deptid) Between " & CodDptoIni & " And " & CodDptoFin & ")) ORDER BY Dept.DeptName, Userinfo.Cardnum"


       Set rpt = New ArpListadoEmpleadosDpto
      End If
      


         rpt.DataControl1.ConnectionString = ConexionEasy
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
    
    
         fPreview.Show 1
         
 Case "LISTADO HORARIOS"
      sql = "SELECT TimeTable.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime FROM TimeTable"
      

         Set rpt = New ArepHorarios
         rpt.DataControl1.ConnectionString = ConexionEasy
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
    
    
         fPreview.Show 1
         
 Case "LISTADO DE EQUIPOS"
      sql = "SELECT FingerClient.* FROM FingerClient"
         Set rpt = New ArepDispositivos
         rpt.DataControl1.ConnectionString = ConexionEasy
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
 Case "REPORTE ASISTENCIA X DIA"

     
      
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaEntradaInicio = FechaHInicio
      FechaEntradaFin = FechaHFinal
                        
      FechaSalidaInicio = FechaHInicio
      FechaSalidaFin = FechaHFinal
      

      
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
        
        If CodEmpleado = 53589 Then
          CodigoH = ""
        End If
        
       

        Me.osProgress2.Min = 0
        Me.osProgress2.Max = DateDiff("d", Me.DTPFechaIni.Value, Me.DTFechaFin.Value)
        Me.osProgress2.Value = 0
        Me.osProgress2.Visible = True
        Contador = 0
        FechaInicial = Me.DTPFechaIni.Value
        Do While FechaInicial <= DTFechaFin.Value
         Me.Caption = "Procesando " & FechaInicial & " Empleado: " & i & " de " & Me.osProgress1.Max & " Codigo: " & CodEmpleado
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
                sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
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
                 Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units, SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                    If Not Me.AdoHorarios.Recordset.EOF Then
                    
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
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
                              
                              FechaEntradaInicio = FechaHInicio
                              FechaEntradaFin = FechaHFinal
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                               FechaSalidaInicio = FechaHInicio
                               FechaSalidaFin = FechaHFinal
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              FechaEntradaInicio = FechaHInicio
                              FechaEntradaFin = FechaHFinal
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              FechaSalidaInicio = FechaHInicio
                              FechaSalidaFin = FechaHFinal
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                           
                           ConfCalcularHorasTrab = True
                           ConfHorasTrabajadas = 0
                           
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
                                    
'                                    '///////////////////////////////////////////////////////////////////////////////////////
'                                    '////////////////BUSCO LA HORAS LABORADAS SEGUN EL HORARIO /////////////////////////////
'                                    '//////////////////////////////////////////////////////////////////////////////////////
'                                    If TipoHorasTrabajada = "Horario" Then
'                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
'                                       ConfCalcularHorasTrab = True
'                                       ConfHorasTrabajadas = (DateDiff("n", Me.AdoHorarios.Recordset("Intime"), Me.AdoHorarios.Recordset("OutTime")) / 60) - RestarAlmuerzo
'                                    End If
            
                                
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
                                   If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                    TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                   Else
                                    TardePermintido = 0
                                   End If
                                   
                                   FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                                   MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                   MinutosTarde = MinutosSalida & ":00" & ":00"
                                   FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                                   FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                                   
                                   FechaEntradaInicio = FechaHInicio
                                   FechaEntradaFin = FechaHFinal
                                  
                    
                                   
                                   sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                         "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                    
                                   
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                   '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
                                           ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                           LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & EInTime))
                                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                           MinutosTarde = MinutosSalida & ":00" & ":00"
                                           LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                                           FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                           FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                                   
            '                       HorasIn = Int(LongitudMinutosIn / 60) & ":" & Int(LongitudMinutosIn Mod 60)
            
            '                       If (CDate(InTime) + CDate(HorasIn)) > CDate(Fecha) Then
            '                        '////SI LA SALIDA ES PARA EL DIA SIGUIENTE PASO PARA EL DIA SIGUIENTE
            '                        FechaHInicio = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & BOutTime & "#"
            '                        FechaHFinal = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & EOutTime & "#" '+ CDate(MinutosTarde) 'Me.DTFechaFin.Value
            ''                        FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
            '                       End If
            
                                   FechaSalidaInicio = FechaHInicio
                                   FechaSalidaFin = FechaHFinal
                                   
                                   SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                               "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                       
                                   
                                                                       '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       If ConfHorasTrabajadas < 0 Then
                                         FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                         ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       End If
                                    End If
                       
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
                        
                        FechaEntradaInicio = FechaHInicio
                        FechaEntradaFin = FechaHFinal
                        
                        FechaSalidaInicio = FechaHInicio
                        FechaSalidaFin = FechaHFinal
                        
                        
                       
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
                          
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                                                           
                          FechaSalidaInicio = FechaHInicio
                          FechaSalidaFin = FechaHFinal
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)= " & Ent & "))"
                        
                          FechaSalidaInicio = FechaHInicio
                          FechaSalidaFin = FechaHFinal
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)= " & Sal & "))"
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
'                                 If JornadaIntercalada = False Then
'                                    If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                       FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                    End If
'                                 End If
                           
                                FechaHInicio = "#" & FechaHInicio & "#"
                                FechaHFinal = "#" & FechaHFinal & "#"
                                
                                SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
'                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) - HorasLaborales) * 3600
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
                                          Me.AdoReportes.Recordset("Campo6") = FechaEntradaInicio
                                          Me.AdoReportes.Recordset("Campo7") = Format(FechaEntradaFin, "dd/mm/yyyy HH:MM")
                                          Me.AdoReportes.Recordset("Campo8") = Format(FechaSalidaInicio, "dd/mm/yyyy HH:MM")
                                          Me.AdoReportes.Recordset("Campo9") = Format(FechaSalidaFin, "dd/mm/yyyy HH:MM")
                                          
                                         
                                         Me.AdoReportes.Recordset.Update
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
                                          Me.AdoBuscaReporte.Recordset("Campo6") = FechaEntradaInicio
                                          Me.AdoBuscaReporte.Recordset("Campo7") = FechaEntradaFin
                                          Me.AdoBuscaReporte.Recordset("Campo8") = FechaSalidaInicio
                                          Me.AdoBuscaReporte.Recordset("Campo9") = FechaSalidaFin
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
      
         
         
        If Me.ChkTodosDptos.Value = 0 Then
            If Me.DBDptoIni.Text = "" Or Me.DBDptoFin.Text = "" Then
             sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin From Reportes ORDER BY Reportes.CampoFecha3, Reportes.CampoNum1,Reportes.CampoFecha1"
             Set rpt = New ArepAsistencia
            Else
             sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin " & _
                   "From Reportes WHERE (((Reportes.Campo3) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Reportes.CampoFecha3, Reportes.Campo3,  Reportes.CampoNum1,Reportes.CampoFecha1"
             Set rpt = New ArepAsistenciaDpto
            End If
        Else
             sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin " & _
                   "From Reportes ORDER BY Reportes.CampoFecha3, Reportes.Campo3,Reportes.CampoNum1,Reportes.CampoFecha1"
             Set rpt = New ArepAsistenciaDpto
        End If
         
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj

 Case "REPORTE LLEGADAS TARDE"
     
      

      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      
      '******************************************************************************
      '//////BUSCO LA CONFIGURACION GENERAL /////////////////////////////////////////
      '*****************************************************************************
       MDIPrimero.DtaEmpresa.Refresh
       If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
         If MDIPrimero.DtaEmpresa.Recordset("RestarToleranciaLlegada") = True Then
            ToleranciaTarde = True
         Else
            ToleranciaTarde = False
         End If
       End If
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      
      If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
'       sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
        sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName, Dept.Deptid FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Dept.Deptid, Checkinout.Userid "
      Else
       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName, Dept.Deptid FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Dept.Deptid) Between " & Me.DBDptoIni.Columns(0).Text & " And " & Me.DBDptoFin.Columns(0).Text & ") AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Dept.Deptid, Checkinout.Userid "
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
        Me.osProgress2.Min = 0
        Me.osProgress2.Max = DateDiff("d", Me.DTPFechaIni.Value, Me.DTFechaFin.Value)
        Me.osProgress2.Value = 0
        Me.osProgress2.Visible = True
        
        Contador = 0
        FechaInicial = Me.DTPFechaIni.Value
        Do While FechaInicial <= DTFechaFin.Value
        
         Me.Caption = "Procesando " & FechaInicial & " Empleado: " & i & " de " & Me.osProgress1.Max
         DoEvents

        
                '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                      "WHERE ((UserShift.Userid)='" & CodEmpleado & "')"
                Me.AdoConsulta.RecordSource = sql
                Me.AdoConsulta.Refresh
                If Not Me.AdoConsulta.Recordset.EOF Then
                  FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                  Ciclo = Me.AdoConsulta.Recordset("Cycles")
'                  Date1 = CDate(FechaInicioH)
'                  Date2 = CDate(Me.DtpFechaINI.Value)
                  Date1 = CDate(FechaInicioH)
                  Date2 = CDate(FechaInicial)
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
        
                
                
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                '////////////BUSCO EL HORARIO PARA ESTE EMPLEADO ////////////////////////////////////////////////////////////////
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                              "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "') ) "
        
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                    Me.osProgress2.Min = 0
                    Me.osProgress2.Value = 0
                    If Not Me.AdoHorarios.Recordset.EOF Then
                      Me.AdoHorarios.Recordset.MoveLast
                      Me.osProgress2.Max = Me.AdoHorarios.Recordset.RecordCount
                      Me.AdoHorarios.Recordset.MoveFirst
                    Else
                       Me.osProgress2.Max = 0
                    End If
                    
                    Me.AdoHorarios.Refresh
                    Do While Not Me.AdoHorarios.Recordset.EOF
                       BInTime = Me.AdoHorarios.Recordset("BIntime")
                       EInTime = Me.AdoHorarios.Recordset("EIntime")
                       InTime = Me.AdoHorarios.Recordset("Intime")
                       
                       
                       BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                       EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                       OutTime = Me.AdoHorarios.Recordset("OutTime")
                       If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                        TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                       Else
                        TardePermintido = 0
                       End If
                       If TardePermintido <= 60 Then
                         MinutosTarde = "00:" & TardePermintido & ":00"
                       End If
                       

                       
                       HoraHorario = CDate(InTime) + CDate(MinutosTarde)
                       
                       FechaIni = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"
                       FechaFin = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & EInTime & "#"
                       
                       FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                       MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                       MinutosTarde = MinutosSalida & ":00" & ":00"
                       FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                       
                       FechaEntradaInicio = FechaHInicio
                       FechaEntradaFin = FechaHFinal
                       
                       '*********************************************************************************************
                       '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                       '*********************************************************************************************
        '               SQL = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
        '                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") "
                        sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & "))"
                        Me.AdoConsulta.RecordSource = sql
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                          Entrada = Me.AdoConsulta.Recordset("CheckTime")
                          HoraEntrada = Format(Entrada, "hh:mm:ss")
                        Else
                          Entrada = "00:00"
                          HoraEntrada = "00:00:00"
                        End If
                        
                    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                    '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                    '/////////////////////////////////////////////////////////////////////////////////////
                    Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                        
                       
'                            If Me.osProgress2.Value < 1 Then
                              If HoraEntrada > HoraHorario Then
                                  Me.AdoReportes.Recordset.AddNew
                                   Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                   Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                   Me.AdoReportes.Recordset("Campo3") = departamento
                                   Me.AdoReportes.Recordset("CampoFecha1") = HoraHorario
                                   Me.AdoReportes.Recordset("CampoFecha2") = HoraEntrada
                                   Me.AdoReportes.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                                   If ToleranciaTarde = True Then
                                     HoraHorario = CDate(InTime) + CDate("00:00:00")
                                     Me.AdoReportes.Recordset("Campo4") = DateDiff("n", HoraHorario, HoraEntrada)
                                   Else
                                     Me.AdoReportes.Recordset("Campo4") = DateDiff("n", HoraHorario, HoraEntrada)
                                   End If
                                   Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                   Me.AdoReportes.Recordset("Campo5") = FechaEntradaInicio
                                   Me.AdoReportes.Recordset("Campo6") = FechaEntradaFin
                                  Me.AdoReportes.Recordset.Update
                               End If
'                              End If
                      End If
                        
                        Me.AdoHorarios.Recordset.MoveNext
                        Me.osProgress2.Value = Me.osProgress2.Value + 1
                    Loop
        
        
        
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
'        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         Me.AdoReportes.Refresh
      
         sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Horario, Reportes.CampoFecha2 AS Entrada, Reportes.Campo4 AS MinutosTarde,Reportes.CampoFecha3 As Fecha,Reportes.Campo5 AS FechaEntradaInicio,Reportes.Campo6 AS FechaEntradaFin  From Reportes ORDER BY Reportes.CampoFecha3, Reportes.Campo3,  Reportes.CampoNum1,Reportes.CampoFecha1"


         Set rpt = New ArepLlegadasTarde
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
      
 Case "REPORTE SALIDA ANTICIPADA"
 
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
'      sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      
     If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
        sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName, Dept.Deptid FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Dept.Deptid, Checkinout.Userid "
      Else
       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName, Dept.Deptid FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((Dept.Deptid) Between " & Me.DBDptoIni.Columns(0).Text & " And " & Me.DBDptoFin.Columns(0).Text & ") AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Dept.Deptid, Checkinout.Userid "
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
        
        Me.osProgress2.Min = 0
        Me.osProgress2.Max = DateDiff("d", Me.DTPFechaIni.Value, Me.DTFechaFin.Value)
        Me.osProgress2.Value = 0
        Me.osProgress2.Visible = True
        
        Contador = 0
        FechaInicial = Me.DTPFechaIni.Value
        Do While FechaInicial <= DTFechaFin.Value
         Me.Caption = "Procesando " & FechaInicial & " Empleado: " & i & " de " & Me.osProgress1.Max
         DoEvents
        

                If CodEmpleado = 760 Then
                  Cod = 1
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
                  Date2 = CDate(Me.DTPFechaIni.Value)
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
        
                
                
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                '////////////BUSCO EL HORARIO PARA ESTE EMPLEADO ////////////////////////////////////////////////////////////////
                '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                              "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
        
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                    Me.osProgress2.Min = 0
                    Me.osProgress2.Value = 0
                    If Not Me.AdoHorarios.Recordset.EOF Then
                      Me.AdoHorarios.Recordset.MoveLast
                      Me.osProgress2.Max = Me.AdoHorarios.Recordset.RecordCount
                      Me.AdoHorarios.Recordset.MoveFirst
                    Else
                       Me.osProgress2.Max = 0
                    End If
                    
                    
                    Do While Not Me.AdoHorarios.Recordset.EOF
                       BInTime = Me.AdoHorarios.Recordset("BIntime")
                       EInTime = Me.AdoHorarios.Recordset("EIntime")
                       InTime = Me.AdoHorarios.Recordset("Intime")
                       LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                       
                       
                       BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                       EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                       OutTime = Me.AdoHorarios.Recordset("OutTime")
                       If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                       TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                       Else
                       TardePermintido = 0
                       End If
                       If TardePermintido <= 60 Then
                         MinutosTarde = "00:" & TardePermintido & ":00"
                       End If
        
                       HoraHorario = CDate(OutTime) '+ CDate(MinutosTarde)
                       
                       HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                              
                       FechaIni = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"
                       FechaFin = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                       
                       FechaEntradaInicio = FechaIni
                       FechaEntradaFin = FechaFin
                       
'                       HorasIn = Int(LongitudMinutosIn / 60) & ":" & Int(LongitudMinutosIn Mod 60)
'                       If (CDate(InTime) + CDate(HorasIn)) > CDate("23:59") Then
'                        '////SI LA SALIDA ES PARA EL DIA SIGUIENTE PASO PARA EL DIA SIGUIENTE
'                        FechaIni = "#" & DateAdd("d", 1, FechaInicial) & " " & BOutTime & "#"
'                        FechaFin = CDate(DateAdd("d", 1, FechaInicial) & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                        FechaFin = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
'                       End If

                       
                       '*********************************************************************************************
                       '///////////////CON ESTA CONSULTA BUSCO LA HORA DE SALIDA///////////////////////////////////
        '               '*********************************************************************************************
        '               SQL = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
        '                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Checkinout.CheckType)=" & sal & "))"
        
                        sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & "))"
                              
        
                        Me.AdoConsulta.RecordSource = sql
                        Me.AdoConsulta.Refresh
                        
                        Entrada = "00:00"
                        If Not Me.AdoConsulta.Recordset.EOF Then
                          Entrada = Me.AdoConsulta.Recordset("CheckTime")
                          HoraEntrada = Format(Entrada, "hh:mm:ss")
                        Else
                          HoraEntrada = Format(Entrada, "hh:mm:ss")
                        End If
                        
                        
                        
                       
                
                        If HoraEntrada < HoraHorario Then
                            Me.AdoReportes.Recordset.AddNew
                             Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                             Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                             Me.AdoReportes.Recordset("Campo3") = departamento
                             Me.AdoReportes.Recordset("CampoFecha1") = HoraHorario
                             Me.AdoReportes.Recordset("CampoFecha2") = HoraEntrada
                             Me.AdoReportes.Recordset("CampoFecha3") = Format(FechaInicial, "dd/mm/yyyy")
                             Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                             Me.AdoReportes.Recordset("Campo4") = DateDiff("n", HoraHorario, HoraEntrada)
                             Me.AdoReportes.Recordset("Campo5") = FechaEntradaInicio
                             Me.AdoReportes.Recordset("Campo6") = FechaEntradaFin
                            Me.AdoReportes.Recordset.Update
                        End If
                        
                        Me.AdoHorarios.Recordset.MoveNext
                        Me.osProgress2.Value = Me.osProgress2.Value + 1
                    Loop
                    
                    
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress1.Value = Me.osProgress1.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
                    
                    
        
        i = i + 1
        Me.osProgress1.Value = i
'        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Horario, Reportes.CampoFecha2 AS Entrada, Reportes.Campo4 AS MinutosTarde,Reportes.CampoFecha3 As Fecha, Reportes.Campo5 AS FechaEntradaInicio, Reportes.Campo6 AS FechaEntradaFin From Reportes ORDER BY Reportes.CampoFecha3, Reportes.Campo3,Reportes.CampoNum1"


         Set rpt = New ArepSalidaAnticipada
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
 
End Select

Me.Frame6.Visible = False
End Sub

Private Sub CmdVerReporte2_Click()
Dim sql As String, CodDptoIni As String, CodDptoFin As String
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
Dim TotalTrabajadas As String, TotalExtras As Date, HorasIn As String, CodigoHorario As String
Dim Horas As String, HoraTarde As String, EntradaAlmuerzo As String, SalidaAlmuerzo As String, EntradaAlmuerzo1 As String, EntradaAlmuerzo2 As String, SalidaAlmuerzo1 As String, SalidaAlmuerzo2 As String, ExcluirSabado As Boolean
Dim SQlEntradaAlmuerzo As String, SqlSalidaAlmuerzo As String, TineJornadas As Boolean
Dim EntradaA As String, SalidaA As String, HoraAlmuerzo As String, DifHoras As Double, DiaExtra As Double, TipoHorasTrabajada As String
Dim Fecha As Date, RestarAlmuerzo As Double, SinHorario As Boolean, TotalHorasTarde As Date, ToleranciaTarde As Boolean
Dim MinutosExtra As Double, MinutosHorasExtra As Double, CantHorarios As Double, SqlIN(6) As String, SqlOut(6) As String, L As Double, HoraInTime(6) As String, HoraOutTime(6) As String, MinutosTardeHorario(6) As String

TieneJornadas = False
Me.osProgress2.Visible = False
CodigoH = ""
Me.AdoDatosEmpresa.Refresh

        If Not IsNull(Me.AdoDatosEmpresa.Recordset("MinutosExtra")) Then
         MinutosExtra = Me.AdoDatosEmpresa.Recordset("MinutosExtra")
        Else
         MinutosExtra = 0
        End If
        
        CantHorarios = 0

      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj

Me.Frame6.Visible = True
Me.Frame6.Top = 1080

Select Case Me.CmbReportes.Text

 Case "REPORTE DE JUSTIFICACION"
 
'      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      

'         sql = "SELECT DISTINCT Userinfo.Userid, Userinfo.Name FROM LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid " & _
               "WHERE (((UserLeave.BeginTime)>=" & FechaIni & ") AND ((UserLeave.EndTime)<=" & FechaFin & "))"

'sql = "SELECT DISTINCT Userinfo.Userid, Userinfo.Name, UserLeave.BeginTime, UserLeave.EndTime, LeaveClass.Classname FROM LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid  " & _
                                          "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserLeave.BeginTime)>=" & FechaIni & ") AND ((UserLeave.EndTime)<=" & FechaFin & "))"

          If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
           If Me.TDBCombo1.Text = "" And Me.DBEmpleado2.Text = "" Then
             If Me.TDBJustificaIni.Text = "" And Me.TDBJustificaFin.Text = "" Then
           
'              sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, LeaveClass.Classname, UserLeave.Whys FROM (LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid " & _
'                    "WHERE (((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & "));"

               sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid " & _
                    "WHERE (((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & "))"
              Else
                sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid WHERE (((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & ") AND ((LeaveClass.Classname) Between '" & Me.TDBJustificaIni.Text & "' And '" & Me.TDBJustificaFin.Text & "'))"
              End If
           
           
           Else
'              sql = "SELECT DISTINCT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, LeaveClass.Classname FROM (LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid " & _
'                    "WHERE (((Userinfo.Userid) In (SELECT DISTINCT Userinfo.Userid FROM (LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((UserLeave.BeginTime) Between " & FechaIni & " And " & FechaFin & "))) Or (Userinfo.Userid) In (SELECT DISTINCT Userinfo.Userid FROM (LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((UserLeave.EndTime) Between " & FechaIni & " And " & FechaFin & "))))) AND (((Userinfo.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "'))"
            If Me.TDBJustificaIni.Text = "" And Me.TDBJustificaFin.Text = "" Then
              sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid " & _
                    "WHERE (((Userinfo.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & "))"
            Else
               sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid " & _
                    "WHERE (((Userinfo.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & ") AND ((LeaveClass.Classname) Between '" & Me.TDBJustificaIni.Text & "' And '" & Me.TDBJustificaFin.Text & "'))"
            End If
           End If
          Else
'             sql = "SELECT DISTINCT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, LeaveClass.Classname FROM (LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid " & _
'                   "WHERE (((Userinfo.Userid) In (SELECT DISTINCT Userinfo.Userid FROM (LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((UserLeave.BeginTime) Between " & FechaIni & " And " & FechaFin & "))) Or (Userinfo.Userid) In (SELECT DISTINCT Userinfo.Userid FROM (LeaveClass INNER JOIN (UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) ON LeaveClass.Classid = UserLeave.LeaveClassid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid WHERE (((UserLeave.EndTime) Between " & FechaIni & " And " & FechaFin & ")))) AND ((Dept.Deptid) Between " & Me.DBDptoIni.Columns(0).Text & " And " & Me.DBDptoFin.Columns(0).Text & "))"
             If Me.TDBJustificaIni.Text = "" And Me.TDBJustificaFin.Text = "" Then
              sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid " & _
                    "WHERE (((Dept.Deptid) Between " & Me.DBDptoIni.Columns(0).Text & " And " & Me.DBDptoFin.Columns(0).Text & ") AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & "))"
             Else
               sql = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid " & _
                    "WHERE (((Dept.Deptid) Between " & Me.DBDptoIni.Columns(0).Text & " And " & Me.DBDptoFin.Columns(0).Text & ") AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & ") AND ((LeaveClass.Classname) Between '" & Me.TDBJustificaIni.Text & "' And '" & Me.TDBJustificaFin.Text & "'))"
             End If
          End If

      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj

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
        
        

      FechaInicial = Me.DTPFechaIni.Value
      Me.AdoReportes.RecordSource = "SELECT Reportes.* From Reportes "
      Me.AdoReportes.Refresh

        Me.AdoEmpleados.Refresh
'        Me.AdoEmpleados.Recordset.MoveFirst
        Do While Not Me.AdoEmpleados.Recordset.EOF
            DoEvents

            CodEmpleado = Me.AdoEmpleados.Recordset("Userid")
                  If Not IsNull(Me.AdoEmpleados.Recordset("Name")) Then
                    NombreEmpleado = Me.AdoEmpleados.Recordset("Name")
                  Else
                    NombreEmpleado = ""
                  End If
                 departamento = Me.AdoEmpleados.Recordset("DeptName")


                    FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
                    

                            Me.AdoReportes.Recordset.AddNew
                             Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                             Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                             Me.AdoReportes.Recordset("Campo3") = departamento
                             Me.AdoReportes.Recordset("CampoFecha1") = Me.AdoEmpleados.Recordset("BeginTime")
                             Me.AdoReportes.Recordset("CampoFecha2") = Me.AdoEmpleados.Recordset("EndTime")
                             Me.AdoReportes.Recordset("Campo4") = Me.AdoEmpleados.Recordset("Whys")
                            Me.AdoReportes.Recordset.Update



            i = i + 1
            Me.osProgress1.Value = i
            Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
            Me.AdoEmpleados.Recordset.MoveNext
        Loop
        
        Me.AdoReportes.Refresh
        

'         sql = "SELECT Reportes.Campo1 AS Userid, Reportes.Campo2 AS Name, Reportes.Campo3 AS DeptName, Reportes.CampoFecha1 AS BeginTime, Reportes.CampoFecha2 AS EndTime, Reportes.Campo4 AS Classname FROM Reportes"
         sql = "SELECT DISTINCT Reportes.Campo1 AS Userid, Reportes.Campo2 AS Name , Reportes.Campo3 AS DeptName FROM Reportes"
         Set rpt = New ArepJustificacion
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
 
 Case "REPORTE HORAS EXTRA SIETE DIAS"
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      

      
      '******************************************************************************
      '//////BUSCO LA CONFIGURACION GENERAL /////////////////////////////////////////
      '*****************************************************************************
       MDIPrimero.DtaEmpresa.Refresh
       If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
         If MDIPrimero.DtaEmpresa.Recordset("RestarToleranciaLlegada") = True Then
            ToleranciaTarde = True
         Else
            ToleranciaTarde = False
         End If
       End If
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
'      SQL = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      Else
       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
             "WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"
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
        TotalHorasExtras = 0
        TotalExtras = 0
        CodigoH = ""
        
        

        
        b = 1
        
          Me.osProgress2.Visible = True
          Me.osProgress2.Max = 6
          Me.osProgress2.Min = 0
          Me.osProgress2.Value = 0
          
          TotalHorasTrabajadas = 0
          TotalTrabajadas = "00:00"
        
          For j = 0 To 6
          
                 If j = 0 Then
                    '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                    sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                          "WHERE ((UserShift.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Me.DTFechaFin.Value, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Me.DTPFechaIni.Value, "YYYY-MM-DD") & "')"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                      Ciclo = Me.AdoConsulta.Recordset("Cycles")
                      Date1 = CDate(FechaInicioH)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                    Else
                      Date1 = CDate(Me.DTPFechaIni.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                      FechaInicioH = CDate(Me.DTPFechaIni)
                    End If
                 Else
                        Date1 = CDate(FechaInicioH)
'                        Date1 = CDate(Me.DtpFechaINI)
                        Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                        DiaInicio = DiaHorario(Date1, Date2, Ciclo)

                End If
                
                Me.Caption = "Procesando " & Date2 & " Empleado: " & i & " de " & Me.osProgress1.Max
                DoEvents
                
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
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                    If Not Me.AdoHorarios.Recordset.EOF Then
                    
                   
                   
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(Date2, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Date2, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                        
                          LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                          If LongitudMinutosIn < 1200 Then  'MENOR A 1400MIN 12 HORAS
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                           
                            ConfCalcularHorasTrab = True
                            ConfHorasTrabajadas = 0
                           
                           SinHorario = True
                           HoraInTime(CantHorarios) = "?"
                           HoraOutTime(CantHorarios) = "?"
                           SqlIN(0) = sql
                           SqlOut(0) = SQlSalida
                           CantHorarios = 1
                      Else
                           TieneJornadas = False
                           SinHorario = False
                           CantHorarios = 0
                           Me.AdoHorarios.Refresh
                           If CodEmpleado = 2 Then
                             CodEmpleado = 2
                           End If
                        
                           Do While Not Me.AdoHorarios.Recordset.EOF
                           
                               BInTime = Me.AdoHorarios.Recordset("BIntime")
                               EInTime = Me.AdoHorarios.Recordset("EIntime")
                               InTime = Me.AdoHorarios.Recordset("Intime")
                               LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                               
'                               Me.AdoHorarios.Recordset.MoveLast
                               
                               BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                               EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                               OutTime = Me.AdoHorarios.Recordset("OutTime")
                               If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                               TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                               Else
                                 TardePermintido = 0
                               End If
                               
                               
                               FechaIn = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               FechaOut = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               
                               FechaHInicio = "#" & FechaIn & " " & BInTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EInTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BInTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                               
                               sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                
                               FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                        
                               FechaEntradaInicio = FechaHInicio
                               FechaEntradaFin = FechaHFinal
                             
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & InTime))
'                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
'
                               
                                ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                        LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                        HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & EInTime))
                                        FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                        MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                        MinutosTarde = MinutosSalida & ":00" & ":00"
                                        LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                              FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                               FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                                                                     
                               SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                
                                '********************************************************************************************
                                '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                '********************************************************************************************
                                
                                FechaSalidaInicio = FechaHInicio
                                FechaSalidaFin = FechaHFinal
                                
                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                CodigoH = Me.AdoHorarios.Recordset("Schid")
                            
                                Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
                                Me.AdoBuscaReporte.Refresh
                                If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                    '/////SI TIENE HORAS EXTRA EN EL HORARIO, SE CAMBIA LA CONFIGURACION GENERAL ////////////
                                    TipoHorasTrabajada = Me.AdoBuscaReporte.Recordset("TipoCalcularHorasTrab")
                                    DiaExtra = DiaSemana(Day(Date2), Month(Date2), Year(Date2))
                                    If DiaExtra = 6 Then
                                           ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrabSab")
                                        ElseIf DiaExtra = 0 Then
                                           ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrabDom")
                                        Else
                                           ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrab")
                                        End If
                                           ConfCalcularHorasTrab = Me.AdoBuscaReporte.Recordset("CalcularHorasTrab")
                                 End If
                                 
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       If ConfHorasTrabajadas < 0 Then
                                         FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                         ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       End If
                                    
                                    End If
                                 
                       If TardePermintido <= 60 Then
                          MinutosTarde = "00:" & TardePermintido & ":00"
                       End If
                       MinutosTardeHorario(CantHorarios) = MinutosTarde
                       HoraInTime(CantHorarios) = InTime
                       HoraOutTime(CantHorarios) = OutTime
                       SqlIN(CantHorarios) = sql
                       SqlOut(CantHorarios) = SQlSalida
                       CantHorarios = CantHorarios + 1
                       Me.AdoHorarios.Recordset.MoveNext
                     Loop

                    
                  End If
                    
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
                       
                       
                      '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                      Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                      Me.AdoBuscaReporte.Refresh
                      If Not Me.AdoBuscaReporte.Recordset.EOF Then
                          CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                          HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                          RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                          RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                          JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                          
                         
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                      End If
                      
                           SinHorario = True
                           HoraInTime(CantHorarios) = InTime
                           HoraOutTime(CantHorarios) = OutTime
                           SqlIN(0) = sql
                           SqlOut(0) = SQlSalida
                           CantHorarios = 1
                      
                    End If

                 For L = 0 To CantHorarios - 1
                        MinutosTarde = MinutosTardeHorario(L)
                        InTime = HoraInTime(L)
                        OutTime = HoraOutTime(L)
                        sql = SqlIN(L)
                        SQlSalida = SqlOut(L)
                        If InTime <> "?" Then
                        HoraHorario = CDate(InTime)
                        End If
                            '*********************************************************************************************
                            '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                            '*********************************************************************************************
                            
                                Entrada = "00:00"
                                HoraEntrada = "00:00"
                                If TieneJornadas = True Then
                                
                                    Me.AdoConsulta.RecordSource = sql
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                      Entrada = Me.AdoConsulta.Recordset("CheckTime")
                                      HoraEntrada = Format(Entrada, "hh:mm:ss")
                                    End If
                               
                                Else
                                    Me.AdoConsulta.RecordSource = sql
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                      Entrada = Me.AdoConsulta.Recordset("CheckTime")
                                      HoraEntrada = Format(Entrada, "hh:mm:ss")
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
                            '///////////////CON ESTA CONSULTA BUSCO LA HORA DE SALIDA///////////////////////////////////
                            '*********************************************************************************************
                                Salida = "00:00"
                                If TieneJornadas = True Then
                                   
                                     '///////////////////////////////CON ESTAS FECHAS BUSCO LA HORA DE SALIDA DE LA JORNADA ///////////////////
                                     
                                     
                                     HoraSalida = CDate(Entrada) + CDate(CInt(HorasLaborales) & ":00:00")
                                     FechaHInicio = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") - CDate(RangoHora1 & ":00")), "mm/dd/yyyy hh:mm:ss")
                                     FechaHFinal = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") + CDate(RangoHora2 & ":00")), "mm/dd/yyyy hh:mm:ss")
'                                     HoraSalida = Format(Date2, "mm/dd/yyyy") & " 23:59:59"
                                     HoraSalida = FechaHFinal
                                     HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                     If JornadaIntercalada = False Then
'                                        If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                           FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                        End If
'                                     End If
                               
                                    FechaHInicio = "#" & FechaHInicio & "#"
                                    FechaHFinal = "#" & FechaHFinal & "#"
                                    
                                    SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                            End If
                            
                      
                            
                            '*********************************************************************************************
                            '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                            '*********************************************************************************************
                            RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                            
                            If Entrada = "00:00" Then
                              Salida = "00:00"
                            End If
                            
                                HorasTrabajadas = 0
                                HoraLaboradas = "00:00"
                                If Salida <> "00:00" Then
                                 If Entrada <> "00:00" Then
            '                      HorasTrabajadas = (DateDiff("h", Entrada, Salida))
                                   HoraLaboradas = ConvertirSegundos((DateDiff("s", Entrada, Salida)), DiaInicio)
                                   HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60) - RestarAlmuerzo  '/////RESTO UNA HORA DE ALMUERZO //////
                                   HoraSalida = Format(Salida, "hh:mm:ss")
                                   TotalTrabajadas = HoraLaboradas + TotalTrabajadas
          
                                 
                                 Else
                                    HorasTrabajadas = 0
                                    HoraLaboradas = "00:00"
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
            '                            RestarAlmuerzo = RestaAlmuerzo(CodigoH)
                                        If TieneJornadas = True Then
                                           If CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) > HorasLaborales Then
                                               HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                               Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                           End If
                                        Else
                                            If ConfCalcularHorasTrab = False Then
                                              If SinHorario = False Then
                                               HorasExtras = (CDbl(((DateDiff("s", HoraSalidaHorario, HoraSalida)) / 3600))) * 3600
                                               Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
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
                                         Horas = "00:00"
                                     End If
                                    Else
                                     HorasExtras = 0
                                     Horas = "00:00"
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
                            
                            HorasExtras = Format(HorasExtras, "##,##0.00")
                            TotalHorasExtras = HorasExtras + TotalHorasExtras
'                            TotalExtras = Horas + TotalExtras
                            
                            
            
                        '--------------------------------------------------------------------------------------------------------------------------------
                        '--------------------------------REPORTE DE LLEGADAS TARDE ----------------------------------------------------------------------
                        '-----------------------------------------------------------------------------------------------------------------------------
                           HoraTarde = "00:00"
                           If InTime <> "?" Then
                                If CDate(HoraEntrada) > (CDate(InTime) + CDate(MinutosTarde)) Then
                                     If ToleranciaTarde = True Then
                                       If InTime <> "?" Then
                                        HoraHorario = CDate(InTime) + CDate("00:00:00")
                                        HorasTarde = DateDiff("S", HoraHorario, HoraEntrada)
                                        HoraTarde = Int(HorasTarde / 3600) & ":" & Int((HorasTarde Mod 3600) / 60)
                                       End If
                                     Else
                                        HoraHorario = CDate(InTime) + CDate(MinutosTarde)
                                        HorasTarde = DateDiff("S", HoraHorario, HoraEntrada)
                                        HoraTarde = Int(HorasTarde / 3600) & ":" & Int((HorasTarde Mod 3600) / 60)
                                     End If
                                 Else
                                    HoraTarde = "00:00"
                                 End If
                           End If

'                            TotalHorasTarde = CDate(HoraTarde) + CDate(TotalHorasTarde)
                            
            
            
                             '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                             '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                             '/////////////////////////////////////////////////////////////////////////////////////
                             Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                             Me.AdoConsulta.Refresh
                             If Not Me.AdoConsulta.Recordset.EOF Then
                            
                                        Select Case j
                                        
                                            Case 0
                                                Me.AdoReportes.Recordset.AddNew
                                                 Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                                 Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                                 Me.AdoReportes.Recordset("Campo3") = departamento
                                                 Me.AdoReportes.Recordset("Campo15") = HoraLaboradas
                                                 Me.AdoReportes.Recordset("Campo22") = Horas
                                                 Me.AdoReportes.Recordset("Campo7") = HoraTarde
                                                 Me.AdoReportes.Recordset.Update
                                                Me.AdoReportes.Refresh
                                             Case 1
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo16") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset("Campo23") = Horas
                                                    Me.AdoBuscaReporte.Recordset("Campo8") = HoraTarde
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                                 
                                             Case 2
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo17") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset("Campo24") = Horas
                                                    Me.AdoBuscaReporte.Recordset("Campo9") = HoraTarde
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 3
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo18") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset("Campo25") = Horas
                                                    Me.AdoBuscaReporte.Recordset("Campo10") = HoraTarde
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 4
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo19") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset("Campo26") = Horas
                                                    Me.AdoBuscaReporte.Recordset("Campo11") = HoraTarde
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 5
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo20") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset("Campo27") = Horas
                                                    Me.AdoBuscaReporte.Recordset("Campo12") = HoraTarde
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                             Case 6
                                                 Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                 Me.AdoBuscaReporte.Refresh
                                                 If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                    Me.AdoBuscaReporte.Recordset("Campo21") = HoraLaboradas
                                                    Me.AdoBuscaReporte.Recordset("Campo28") = Horas
                                                    Me.AdoBuscaReporte.Recordset("Campo5") = Format(TotalTrabajadas, "hh:mm")
                                                    Me.AdoBuscaReporte.Recordset("Campo6") = Format(TotalExtras, "hh:mm")
                                                    Me.AdoBuscaReporte.Recordset("Campo13") = Format(HoraTarde, "hh:mm")
                                                    Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                    Me.AdoBuscaReporte.Recordset.Update
                                                 End If
                                        End Select
                             End If
                Next
                Me.osProgress2.Value = j + 1
        
           Next
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
        Me.AdoBuscaReporte.Refresh
      Loop
      
         
      
         sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.Campo22 AS Dia1, Reportes.Campo23 AS Dia2, Reportes.Campo24 AS Dia3, Reportes.Campo25 AS Dia4, Reportes.Campo26 AS Dia5, Reportes.Campo27 AS Dia6, Reportes.Campo28 AS Dia7, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7, Reportes.Campo11 AS TotalHoras From Reportes ORDER BY Reportes.Campo3,Reportes.Campo1"
           
         
         Set rpt = New ArepHorasExtraSiete
'         Set rpt = New ArepExtraSiete
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj

 Case "REPORTE ASISTENCIA Y AUSENCIA X DIA"
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaEntradaInicio = FechaHInicio
      FechaEntradaFin = FechaHFinal
                        
      FechaSalidaInicio = FechaHInicio
      FechaSalidaFin = FechaHFinal
      

      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      If Me.TDBCombo1.Text = "" And Me.DBEmpleado2.Text = "" Then
        If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
    '        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
             sql = "SELECT DISTINCT Checkinout.Userid From Checkinout"
        Else
             sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM Dept INNER JOIN (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) ON Dept.Deptid = Userinfo.Deptid WHERE (((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "'))"
        End If
      Else
        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
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
         
         Entrada = "00:00"
         Salida = "00:00"
         EntradaA = "00:00"
         SalidaA = "00:00"
         HorasTrabajadas = "00:00"
         Horas = "00:00"
         
            '********************************************************************************************
            '///////////////CON ESTA CONSULTA BUSCO LOS DATOS DE CONFIGURACION //////////////////////////
            '********************************************************************************************
                MDIPrimero.DtaEmpresa.Refresh
                If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
'                  FechaInicial = Me.DTPFechaIni.Value
                   FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
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
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                          
                          LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                         If LongitudMinutosIn < 1200 Then  'Menor a 1400  12horas
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
                              FechaEntradaInicio = FechaHInicio
                              FechaEntradaFin = FechaHFinal
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              FechaSalidaInicio = FechaHInicio
                              FechaSalidaFin = FechaHFinal
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              
                              FechaEntradaInicio = FechaHInicio
                              FechaEntradaFin = FechaHFinal
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              FechaSalidaInicio = FechaHInicio
                              FechaSalidaFin = FechaHFinal
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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

                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                CodigoH = Me.AdoHorarios.Recordset("Schid")
                                 '*******************************************************************************************************************
                                 '*********************************BUSCO EL HORARIO DE ALMUERZO *****************************************************
                                 '*******************************************************************************************************************
                                 Me.AdoHorarioAlmuerzo.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                                    
                                    
                                    Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                       
                       
'                                    Me.AdoHorarios.Recordset.MoveLast
                                    
                                    BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                                    EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                                    OutTime = Me.AdoHorarios.Recordset("OutTime")
                                    LongitudMinutosOut = Me.AdoHorarios.Recordset("Longtime")
                                    If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                     TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                     Else
                                      TardePermintido = 0
                                    End If
                                    
                                    FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                                    MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                    MinutosTarde = MinutosSalida & ":00" & ":00"
                                    FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                                    FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
     
                                    FechaEntradaInicio = FechaHInicio
                                    FechaEntradaFin = FechaHFinal
     
                                   sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                         "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                    
                                   
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                   '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'                                    HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
'                                    FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                                    MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                                    MinutosTarde = MinutosSalida & ":00" & ":00"
'                                    FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                                    FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                                           ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                           
                                           LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & EInTime))
                                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                           MinutosTarde = MinutosSalida & ":00" & ":00"
                                           LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                                          FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                           FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                       
                                           FechaSalidaInicio = FechaHInicio
                                           FechaSalidaFin = FechaHFinal
                       
                                    SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                          
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
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
                                     SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                     
                                     
                                     FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                                     FechaHFinal = CDate(FechaInicial)
                                     FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                                     SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                 
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
                          
                         
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          FechaSalidaInicio = FechaHInicio
                          FechaSalidaFin = FechaHFinal
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          FechaSalidaInicio = FechaHInicio
                          FechaSalidaFin = FechaHFinal
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
                            SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                            
                            
                            FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                            FechaHFinal = CDate(FechaInicial)
                            FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                            SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
                            If Not SQlEntradaAlmuerzo = "" Then
                                
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
                            If Not SqlSalidaAlmuerzo = "" Then
                                
                                Me.AdoConsulta.RecordSource = SqlSalidaAlmuerzo
                                Me.AdoConsulta.Refresh
                                If Not Me.AdoConsulta.Recordset.EOF Then
                                  SalidaA = Me.AdoConsulta.Recordset("CheckTime")
                                End If
                                
                                If ExcluirSabado = True Then
                                  If DiaInicio = 6 Then
                                    EntradaA = "00:00"
                                    SalidaA = "00:00"
                                  End If
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
'                                 HoraSalida = Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59"
                                 HoraSalida = FechaHFinal
                                 HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                 If JornadaIntercalada = False Then
'                                    If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                       FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                    End If
'                                 End If
                           
                                FechaHInicio = "#" & FechaHInicio & "#"
                                FechaHFinal = "#" & FechaHFinal & "#"
                                
                                SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                                   If InTime <> "?" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
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
                                    
                                    HoraAlmuerzo = "00:00"
                                    If Not EntradaA = "00:00" Then
                                     If Not SalidaA = "00:00" Then
            '                            DifHoras = (CDbl(((DateDiff("s", EntradaA, SalidaA)) / 3600) - 1)) * 3600
                                        DifHoras = DateDiff("s", EntradaA, SalidaA)
                                        HoraAlmuerzo = Int(DifHoras / 3600) & ":" & Int((DifHoras Mod 3600) / 60)
                                     End If
                                    End If
         
                                    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                                    '/////////////////////////////////////////////////////////////////////////////////////
                                    Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                    FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
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
                                             Me.AdoReportes.Recordset("CampoFecha6") = Format(EntradaA, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("CampoFecha7") = Format(SalidaA, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("CampoFecha8") = Format(HoraAlmuerzo, "hh:mm:ss")
                                             
                                             Me.AdoReportes.Recordset("Campo6") = FechaEntradaInicio
                                             Me.AdoReportes.Recordset("Campo7") = FechaEntradaFin
                                             Me.AdoReportes.Recordset("Campo8") = FechaSalidaInicio
                                             Me.AdoReportes.Recordset("Campo9") = FechaSalidaFin
                                            Me.AdoReportes.Recordset.Update
                                    End If
                Next
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         Me.AdoReportes.Refresh
      
         
         
   If Me.ChkTodosDptos.Value = 0 Then
         If Me.DBDptoIni.Text = "" Or Me.DBDptoFin.Text = "" Then
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo,Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin From Reportes ORDER BY Reportes.CampoFecha3, Reportes.CampoNum1,Reportes.CampoFecha1"
                Set rpt = New ArepAsistenciaAusencia
                rpt.GroupHeader2.Visible = False
                rpt.GroupFooter2.Visible = False
         Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo , Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin " & _
                      "From Reportes WHERE (((Reportes.Campo3) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Reportes.CampoFecha3, Reportes.Campo3,  Reportes.CampoNum1,Reportes.CampoFecha1"
                Set rpt = New ArepAsistenciaAusencia
                rpt.GroupHeader2.Visible = True
                rpt.GroupFooter2.Visible = True
                rpt.GroupFooter1.Visible = False
        End If
    Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin  " & _
                      "From Reportes ORDER BY Reportes.CampoFecha3, Reportes.Campo3,  Reportes.CampoNum1,Reportes.CampoFecha1"
                Set rpt = New ArepAsistenciaAusencia
                rpt.GroupHeader2.Visible = True
                rpt.GroupFooter2.Visible = True
    End If
         
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
 



 Case "REPORTE DETALLE ASISTENCIA"
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      


      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      If Me.TDBCombo1.Text = "" And Me.DBEmpleado2.Text = "" Then
        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
      Else
        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"

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
        
        If CodEmpleado = 700 Then
         Contador = 0
        End If
        
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
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                          
                          LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                          If LongitudMinutosIn < 1200 Then  'Menor a 1400  12horas
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                           
                            ConfCalcularHorasTrab = True
                            ConfHorasTrabajadas = 0
                            
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
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
                                           
                                           FechaEntradaInicio = FechaHInicio
                                           FechaEntradaFin = FechaHFinal
                            
                                           
                                           sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                            
                                       
                                           Dim HorasTotal As String
                                           '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                           '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                                           '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                           ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                           LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & EInTime))
                                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                           MinutosTarde = MinutosSalida & ":00" & ":00"
                                           LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                           HorasTotal = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                                           FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                           FechaHFinal = "#" & Format(Mid(HorasTotal, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                                           
                                          
                                           SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                       "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                               
                               
                                          FechaSalidaInicio = FechaHInicio
                                          FechaSalidaFin = FechaHFinal
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
                          
                         
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          FechaSalidaInicio = FechaHInicio
                          FechaSalidaFin = FechaHFinal
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          FechaSalidaInicio = FechaHInicio
                          FechaSalidaFin = FechaHFinal
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
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
                            '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                            '*********************************************************************************************
                    
                            Entrada = "00:00"
                            If TieneJornadas = True Then
                            
                                Me.AdoConsulta.RecordSource = sql
                                Me.AdoConsulta.Refresh
                                If JornadaIntercalada = True Then
                                  Me.AdoConsulta.Recordset.MoveLast
                                End If
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
'                                 HoraSalida = Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59"
                                 HoraSalida = FechaHFinal
                                 HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                 If JornadaIntercalada = False Then
'                                    If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                       FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                    End If
'                                 End If
                           
                                FechaHInicio = "#" & FechaHInicio & "#"
                                FechaHFinal = "#" & FechaHFinal & "#"
                                
                                SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                            "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
        
                           
                                Me.AdoConsulta.RecordSource = SQlSalida
                                Me.AdoConsulta.Refresh
                                If Not Me.AdoConsulta.Recordset.EOF Then
                                    Me.AdoConsulta.Recordset.MoveLast
                                    Salida = Me.AdoConsulta.Recordset("CheckTime")
                                ElseIf JornadaIntercalada = True Then
                                  '//////////////SI LA JORNADA ES INTERCALADA Y NO TIENE REGISTRO DE SALIDA /////////////////////////
                                  '//////////////HAGO CERO LA ENTRADA ///////////////////////////////////////////////////////
        '                            Entrada = "00:00"
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
'                          Salida = "00:00"
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

                    
                            '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                            '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                            '/////////////////////////////////////////////////////////////////////////////////////
                            Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                            Me.AdoConsulta.Refresh
                            If Not Me.AdoConsulta.Recordset.EOF Then
                            
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
                                     Me.AdoReportes.Recordset("Campo6") = Format(FechaEntradaInicio, "dd/mm/yyyy HH:MM")
                                     Me.AdoReportes.Recordset("Campo7") = Format(FechaEntradaFin, "dd/mm/yyyy HH:MM")
                                     Me.AdoReportes.Recordset("Campo8") = Format(FechaSalidaInicio, "dd/mm/yyyy HH:MM")
                                     Me.AdoReportes.Recordset("Campo9") = Format(FechaSalidaFin, "dd/mm/yyyy HH:MM")
                                    Me.AdoReportes.Recordset.Update
                            End If
            Next
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         Me.AdoReportes.Refresh
      
         
         

          sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin From Reportes ORDER BY Reportes.CampoNum1, Reportes.CampoFecha3"
          Set rpt = New ArepDetalleAsistencia2

         
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj



End Select

Me.Frame6.Visible = False
End Sub

Private Sub CmdVerReporte3_Click()
Dim sql As String, CodDptoIni As String, CodDptoFin As String
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
Dim TotalTrabajadas As String, TotalExtras As Date, HorasIn As String, CodigoHorario As String
Dim Horas As String, HoraTarde As String, EntradaAlmuerzo As String, SalidaAlmuerzo As String, EntradaAlmuerzo1 As String, EntradaAlmuerzo2 As String, SalidaAlmuerzo1 As String, SalidaAlmuerzo2 As String, ExcluirSabado As Boolean
Dim SQlEntradaAlmuerzo As String, SqlSalidaAlmuerzo As String, TineJornadas As Boolean
Dim EntradaA As String, SalidaA As String, HoraAlmuerzo As String, DifHoras As Double, DiaExtra As Double, TipoHorasTrabajada As String
Dim Fecha As Date, RestarAlmuerzo As Double, SinHorario As Boolean, TotalHorasTarde As String, ToleranciaTarde As Boolean
Dim MinutosExtra As Double, MinutosHorasExtra As Double, CantHorarios As Double, SqlIN(6) As String, SqlOut(6) As String, L As Double, HoraInTime(6) As String, HoraOutTime(6) As String, MinutosTardeHorario(6) As String
Dim TotalTarde(6) As String, HorasLabReal As Double

TieneJornadas = False
Me.osProgress2.Visible = False
CodigoH = ""
Me.AdoDatosEmpresa.Refresh

        If Not IsNull(Me.AdoDatosEmpresa.Recordset("MinutosExtra")) Then
         MinutosExtra = Me.AdoDatosEmpresa.Recordset("MinutosExtra")
        Else
         MinutosExtra = 0
        End If
        
        CantHorarios = 0

      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


Me.Frame6.Visible = True
Me.Frame6.Top = 1080

Select Case Me.CmbReportes.Text
  
  




  Case "REPORTE DE ENTRADAS Y SALIDAS"
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaEntradaInicio = FechaHInicio
      FechaEntradaFin = FechaHFinal
                        
      FechaSalidaInicio = FechaHInicio
      FechaSalidaFin = FechaHFinal
      
    

      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      If Me.TDBCombo1.Text = "" And Me.DBEmpleado2.Text = "" Then
        If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
            sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
'             sql = "SELECT DISTINCT Checkinout.Userid From Checkinout"
        Else
             sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM Dept INNER JOIN (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) ON Dept.Deptid = Userinfo.Deptid WHERE (((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "'))"
        End If
      Else
        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
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
         
         Entrada = "00:00"
         Salida = "00:00"
         EntradaA = "00:00"
         SalidaA = "00:00"
         HorasTrabajadas = "00:00"
         Horas = "00:00"
         
            '********************************************************************************************
            '///////////////CON ESTA CONSULTA BUSCO LOS DATOS DE CONFIGURACION //////////////////////////
            '********************************************************************************************
                MDIPrimero.DtaEmpresa.Refresh
                If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
'                  FechaInicial = Me.DTPFechaIni.Value
                   FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
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
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                          
                          LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                         If LongitudMinutosIn < 1200 Then  'Menor a 1400  12horas
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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

                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                CodigoH = Me.AdoHorarios.Recordset("Schid")
                                 '*******************************************************************************************************************
                                 '*********************************BUSCO EL HORARIO DE ALMUERZO *****************************************************
                                 '*******************************************************************************************************************
                                 Me.AdoHorarioAlmuerzo.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                                    
                                    
                                    Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                                    
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       ConfHorasTrabajadas = (DateDiff("n", Me.AdoHorarios.Recordset("Intime"), Me.AdoHorarios.Recordset("OutTime")) / 60) - RestarAlmuerzo
                                    End If
                    
                                    TieneJornadas = False
                                    
                                      BInTime = Me.AdoHorarios.Recordset("BIntime")
                                      EInTime = Me.AdoHorarios.Recordset("EIntime")
                                      InTime = Me.AdoHorarios.Recordset("Intime")
                                      LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                       
                       
'                                    Me.AdoHorarios.Recordset.MoveLast
                                    
                                    BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                                    EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                                    OutTime = Me.AdoHorarios.Recordset("OutTime")
                                    LongitudMinutosOut = Me.AdoHorarios.Recordset("Longtime")
                                    If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                     TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                    Else
                                     TardePermintido = 0
                                    End If
                                    
                                    FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                                    MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                    MinutosTarde = MinutosSalida & ":00" & ":00"
                                    FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                                    FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
     
                                    FechaEntradaInicio = FechaHInicio
                                    FechaEntradaFin = FechaHFinal
                                   
                                   sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                         "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                    
                                   
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                   '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'                                    HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
'                                    FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                                    MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                                    MinutosTarde = MinutosSalida & ":00" & ":00"
'                                    FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                                    FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                       
                                           ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                           LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & EInTime))
                                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                           MinutosTarde = MinutosSalida & ":00" & ":00"
                                           LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                                          FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                           FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                       
                                           FechaSalidaInicio = FechaHInicio
                                           FechaSalidaFin = FechaHFinal
                                    
                                    SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                          
                                                                         '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
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
                                     SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                     
                                     
                                     FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                                     FechaHFinal = CDate(FechaInicial)
                                     FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                                     SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                 
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
                          
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          FechaSalidaInicio = FechaHInicio
                          FechaSalidaFin = FechaHFinal
                                           
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          
                          FechaEntradaInicio = FechaHInicio
                          FechaEntradaFin = FechaHFinal
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                        
                         FechaSalidaInicio = FechaHInicio
                         FechaSalidaFin = FechaHFinal
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
                            SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                            
                            
                            FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                            FechaHFinal = CDate(FechaInicial)
                            FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                            SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
                            If Not SQlEntradaAlmuerzo = "" Then
                                
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
                            If Not SqlSalidaAlmuerzo = "" Then
                                
                                Me.AdoConsulta.RecordSource = SqlSalidaAlmuerzo
                                Me.AdoConsulta.Refresh
                                If Not Me.AdoConsulta.Recordset.EOF Then
                                  SalidaA = Me.AdoConsulta.Recordset("CheckTime")
                                End If
                                
                                If ExcluirSabado = True Then
                                  If DiaInicio = 6 Then
                                    EntradaA = "00:00"
                                    SalidaA = "00:00"
                                  End If
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
'                                 HoraSalida = Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59"
                                 HoraSalida = FechaHFinal
                                 HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                 If JornadaIntercalada = False Then
'                                    If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                       FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                    End If
'                                 End If
                           
                                FechaHInicio = "#" & FechaHInicio & "#"
                                FechaHFinal = "#" & FechaHFinal & "#"
                                
                                SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                                   If InTime <> "?" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
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
                                    
                                    HoraAlmuerzo = "00:00"
                                    If Not EntradaA = "00:00" Then
                                     If Not SalidaA = "00:00" Then
            '                            DifHoras = (CDbl(((DateDiff("s", EntradaA, SalidaA)) / 3600) - 1)) * 3600
                                        DifHoras = DateDiff("s", EntradaA, SalidaA)
                                        HoraAlmuerzo = Int(DifHoras / 3600) & ":" & Int((DifHoras Mod 3600) / 60)
                                     End If
                                    End If
         
                                    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                                    '/////////////////////////////////////////////////////////////////////////////////////
                                    Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                    FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
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
                                             Me.AdoReportes.Recordset("CampoFecha6") = Format(EntradaA, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("CampoFecha7") = Format(SalidaA, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("CampoFecha8") = Format(HoraAlmuerzo, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("Campo6") = FechaEntradaInicio
                                             Me.AdoReportes.Recordset("Campo7") = FechaEntradaFin
                                             Me.AdoReportes.Recordset("Campo8") = FechaSalidaInicio
                                             Me.AdoReportes.Recordset("Campo9") = FechaSalidaFin
                                            Me.AdoReportes.Recordset.Update
                                    End If
                Next
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         Me.AdoReportes.Refresh
      
         
         
   If Me.ChkTodosDptos.Value = 0 Then
         If Me.DBDptoIni.Text = "" Or Me.DBDptoFin.Text = "" Then
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin From Reportes ORDER BY Reportes.Campo1,Reportes.CampoFecha3"
                Set rpt = New ArepAsistenciaRegistros
''               rpt.GroupHeader2.Visible = False
'                rpt.GroupFooter2.Visible = False
         Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin  " & _
                      "From Reportes WHERE (((Reportes.Campo3) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Reportes.Campo1, Reportes.CampoNum1,Reportes.CampoFecha3"
                Set rpt = New ArepAsistenciaRegistros
'                rpt.GroupHeader2.Visible = True
'                rpt.GroupFooter2.Visible = True
'                rpt.GroupFooter1.Visible = False
        End If
    Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin  " & _
                      "From Reportes ORDER BY Reportes.Campo1, Reportes.CampoFecha1"
                Set rpt = New ArepAsistenciaRegistros
'                rpt.GroupHeader2.Visible = True
'                rpt.GroupFooter2.Visible = True
    End If
         
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj


         Case "REPORTE HORAS LAB EXTRA SIETE DIAS"
              FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
              FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
              
        
              
              '******************************************************************************
              '//////BUSCO LA CONFIGURACION GENERAL /////////////////////////////////////////
              '*****************************************************************************
               MDIPrimero.DtaEmpresa.Refresh
               If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
                 If MDIPrimero.DtaEmpresa.Recordset("RestarToleranciaLlegada") = True Then
                    ToleranciaTarde = True
                 Else
                    ToleranciaTarde = False
                 End If
               End If
              
              '*********************************************************************************
              '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
              '*********************************************************************************
               rs.Open "DELETE FROM [Reportes] ", ConexionReloj
        
        
              '****************************************************************************************************************************
              '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
              '***************************************************************************************************************************
        '      SQL = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
              If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
                sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
              Else
               sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
                     "WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"
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
                TotalHorasExtras = 0
                TotalExtras = 0
                CodigoH = ""
                
                
        
                
                b = 1
                
                  Me.osProgress2.Visible = True
                  Me.osProgress2.Max = 6
                  Me.osProgress2.Min = 0
                  Me.osProgress2.Value = 0
                  
                  TotalHorasTrabajadas = 0
                  TotalTrabajadas = "00:00"
                
                  For j = 0 To 6
                  
                         If j = 0 Then
                            '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                            sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                                  "WHERE ((UserShift.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Me.DTFechaFin.Value, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Me.DTPFechaIni.Value, "YYYY-MM-DD") & "')"
                            Me.AdoConsulta.RecordSource = sql
                            Me.AdoConsulta.Refresh
                            If Not Me.AdoConsulta.Recordset.EOF Then
                              FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                              Ciclo = Me.AdoConsulta.Recordset("Cycles")
                              Date1 = CDate(FechaInicioH)
                              Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                              DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                              FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                            Else
                              Date1 = CDate(Me.DTPFechaIni.Value)
                              Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                              DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                              FechaInicioH = CDate(Me.DTPFechaIni.Value)
                            End If
                         Else
                                Date1 = CDate(FechaInicioH)
'                                Date1 = CDate(Me.DtpFechaINI)
                                Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                                DiaInicio = DiaHorario(Date1, Date2, Ciclo)
        
                        End If
                        
                        Me.Caption = "Procesando " & Date2 & " Empleado: " & CodEmpleado & "  " & i & " de " & Me.osProgress1.Max
                        DoEvents
                        
                        
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
                                                       "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                         Me.AdoHorarios.Refresh
                      
                      '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                            If Not Me.AdoHorarios.Recordset.EOF Then
                            
                           
                           
                              Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                            "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(Date2, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Date2, "YYYY-MM-DD") & "')) "
                              Me.AdoHorarios.Refresh
                              If Me.AdoHorarios.Recordset.EOF Then
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                                '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                              "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                                Me.AdoHorarios.Refresh
                                
                                  LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                                  
                               BInTime = "?"
                               EInTime = "?"
                               InTime = "?"
                               
                              
                               BOutTime = "?"
                               EOutTime = "?"
                               OutTime = "?"
                                   
                                   
                                  If LongitudMinutosIn < 1200 Then  'MENOR A 1400MIN 12 HORAS
                                     '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                                      FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 00:00#"
                                      FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                                      
                                      sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                                    
                                      SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                                  Else
                                      FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                                      FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                                     '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                                      sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                                    
                                      SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                                   End If
                                       
                                       ConfCalcularHorasTrab = True
                                       ConfHorasTrabajadas = 0
                                   
                                   SinHorario = True
                                   HoraInTime(0) = "?"
                                   HoraOutTime(0) = "?"
                                   SqlIN(0) = sql
                                   SqlOut(0) = SQlSalida
                                   CantHorarios = 1
                              Else
                                   TieneJornadas = False
                                   SinHorario = False
                                   CantHorarios = 0
                                   Me.AdoHorarios.Refresh
                                
                                   Do While Not Me.AdoHorarios.Recordset.EOF
                                   
                                       BInTime = Me.AdoHorarios.Recordset("BIntime")
                                       EInTime = Me.AdoHorarios.Recordset("EIntime")
                                       InTime = Me.AdoHorarios.Recordset("Intime")
                                       LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                                       
        '                               Me.AdoHorarios.Recordset.MoveLast
                                       
                                       BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                                       EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                                       OutTime = Me.AdoHorarios.Recordset("OutTime")
                                       If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                        TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                       Else
                                         TardePermintido = 0
                                       End If
                                       
                                       
                                       FechaIn = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                                       FechaOut = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                                       
                                       FechaHInicio = "#" & FechaIn & " " & BInTime & "#"
                '                       FechaHFinal = "#" & FechaOut & " " & EInTime & "#"
                                       MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                       MinutosTarde = MinutosSalida & ":00" & ":00"
                                       FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BInTime) + CDate(MinutosTarde)
                                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                                       
                                       sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                        
                                       FechaEntradaInicio = FechaHInicio
                                       FechaEntradaFin = FechaHFinal
                        
                                       FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
                '                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
                                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                       MinutosTarde = MinutosSalida & ":00" & ":00"
                                       FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
                                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                                  
                                   
            
'                                       HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & InTime))
'                                       FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                                       MinutosTarde = MinutosSalida & ":00" & ":00"
'                                       FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                                       FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                                       
                               ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                        LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                        HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & EInTime))
                                        FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                        MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                        MinutosTarde = MinutosSalida & ":00" & ":00"
                                        LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                              FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                               FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                                       
                                       SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                        
                                        '********************************************************************************************
                                        '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                        '********************************************************************************************
                                        
                                         FechaSalidaInicio = FechaHInicio
                                         FechaSalidaFin = FechaHFinal
                                        
                                        CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                        CodigoH = Me.AdoHorarios.Recordset("Schid")
                                    
                                        Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
                                        Me.AdoBuscaReporte.Refresh
                                        If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                            '/////SI TIENE HORAS EXTRA EN EL HORARIO, SE CAMBIA LA CONFIGURACION GENERAL ////////////
                                            TipoHorasTrabajada = Me.AdoBuscaReporte.Recordset("TipoCalcularHorasTrab")
                                            DiaExtra = DiaSemana(Day(Date2), Month(Date2), Year(Date2))
                                            If DiaExtra = 6 Then
                                                   ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrabSab")
                                                ElseIf DiaExtra = 0 Then
                                                   ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrabDom")
                                                Else
                                                   ConfHorasTrabajadas = Me.AdoBuscaReporte.Recordset("HorasTrab")
                                             End If
                                                   ConfCalcularHorasTrab = Me.AdoBuscaReporte.Recordset("CalcularHorasTrab")
                                        End If
                                        

                                    
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
                                        
                               If TardePermintido <= 60 Then
                                  MinutosTarde = "00:" & TardePermintido & ":00"
                               End If
                               MinutosTardeHorario(CantHorarios) = MinutosTarde
                               HoraInTime(CantHorarios) = InTime
                               HoraOutTime(CantHorarios) = OutTime
                               SqlIN(CantHorarios) = sql
                               SqlOut(CantHorarios) = SQlSalida
                               CantHorarios = CantHorarios + 1
                               Me.AdoHorarios.Recordset.MoveNext
                             Loop
        
                            
                          End If
                            
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
                               
                               
                              '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                              Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                              Me.AdoBuscaReporte.Refresh
                              If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                  CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                                  HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                                  RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                                  RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                                  JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                                  
                                 
                                  
                                  sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                                
                                  SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                                  
                                  TieneJornadas = True
                             
                              Else
                              
                                  TieneJornadas = False
                                  sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                                
                                  SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                              End If
                              
                                   SinHorario = True
                                   HoraInTime(CantHorarios) = InTime
                                   HoraOutTime(CantHorarios) = OutTime
                                   SqlIN(0) = sql
                                   SqlOut(0) = SQlSalida
                                   CantHorarios = 1
                              
                            End If

                         For L = 0 To CantHorarios - 1
                         
                                    If CodEmpleado = "3430" Then
                                      CodEmpleado = "3430"
                                    End If
                         
                                MinutosTarde = MinutosTardeHorario(L)
                                InTime = HoraInTime(L)
                                OutTime = HoraOutTime(L)
                             
                                sql = SqlIN(L)
                                SQlSalida = SqlOut(L)

                                    '*********************************************************************************************
                                    '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                                    '*********************************************************************************************
                                    
                                        Entrada = "00:00"
                                        HoraEntrada = "00:00"
                                        If TieneJornadas = True Then
                                        
                                            Me.AdoConsulta.RecordSource = sql
                                            Me.AdoConsulta.Refresh
                                            If Not Me.AdoConsulta.Recordset.EOF Then
                                              Entrada = Me.AdoConsulta.Recordset("CheckTime")
                                              HoraEntrada = Format(Entrada, "hh:mm:ss")
                                            End If
                                       
                                        Else
                                            Me.AdoConsulta.RecordSource = sql
                                            Me.AdoConsulta.Refresh
                                            If Not Me.AdoConsulta.Recordset.EOF Then
                                              Entrada = Me.AdoConsulta.Recordset("CheckTime")
                                              HoraEntrada = Format(Entrada, "hh:mm:ss")
                                            End If
                                        End If
                                        
                                        
                                        '*********************************************************************************************
                                        '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                                        '*********************************************************************************************
                                      If Entrada <> "00:00" Then
                                        If ConfCalcularHorasTrab = True Then
                                            If TipoHorasTrabajada = "HorasTrab" Then
                                              If InTime <> "?" Then
                                               If InTime > Format(Entrada, "hh:mm") Then
                                                  Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                               End If
                                              End If
                                            End If
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
'                                             HoraSalida = Format(Date2, "mm/dd/yyyy") & " 23:59:59"
                                             HoraSalida = FechaHFinal
                                             HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                             If JornadaIntercalada = False Then
'                                                If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                                   FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                                End If
'                                             End If
                                       
                                            FechaHInicio = "#" & FechaHInicio & "#"
                                            FechaHFinal = "#" & FechaHFinal & "#"
                                            
                                            SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                                    
                              
                                     If CodEmpleado = 4 Then
                                       CodEmpleado = 4
                                     End If
                                    '*********************************************************************************************
                                    '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                                    '*********************************************************************************************
                                    RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                    
                                    If Entrada = "00:00" Then
                                      Salida = "00:00"
                                    End If
                                    
                                        HorasTrabajadas = 0
                                        HoraLaboradas = "00:00"
                                        HorasLabReal = 0
                                                                           
                                        If Salida <> "00:00" Then
                                         If Entrada <> "00:00" Then
'                                           HoraLaboradas = ConvertirSegundos((DateDiff("s", Entrada, Salida)))
                                            
                                            If ConfHorasTrabajadas <> 0 Then  ''''EN PRUBA NAVIC
                                               HoraLaboradas = ConvertirSegundos((DateDiff("s", Entrada, Salida)), DiaInicio)
                                               HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60) - RestarAlmuerzo  '/////RESTO UNA HORA DE ALMUERZO //////
                                            End If

                                            HoraSalida = Format(Salida, "hh:mm:ss")
                                           TotalTrabajadas = HoraLaboradas + TotalTrabajadas
                    
                                         
                                         Else
                                            HorasTrabajadas = 0
                                            HoraLaboradas = "00:00"
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
                    '                            RestarAlmuerzo = RestaAlmuerzo(CodigoH)
                                                If TieneJornadas = True Then
                                                   If CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) > HorasLaborales Then
                                                       HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                                       Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                                   End If
                                                Else
                                                    If ConfCalcularHorasTrab = False Then
                                                      If SinHorario = False Then
                                                       If CDate(HoraSalida) > CDate(HoraSalidaHorario) Then
                                                         HorasExtras = (CDbl(((DateDiff("s", HoraSalidaHorario, HoraSalida)) / 3600))) * 3600
                                                         Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                                       Else
                                                         HorasExtras = 0
                                                         Horas = "0:00"
                                                       End If
                                                      Else
                                                       HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600))) * 3600
                                                       Horas = Int(HorasExtras / 3600) - RestarAlmuerzo & ":" & Int((HorasExtras Mod 3600) / 60)
                                                      End If
                                                     
                                                       
                                                    ElseIf CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) > ConfHorasTrabajadas Then
                                                       
                                                       HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) - ConfHorasTrabajadas) * 3600
                                                       Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                                    End If
                                                End If
                                                
                                                
                                             Else
                                                 HorasExtras = 0
                                                 Horas = "00:00"
                                             End If
                                            Else
                                             HorasExtras = 0
                                             Horas = "00:00"
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
                                    
                                    HorasExtras = Format(HorasExtras, "##,##0.0000")
                                    TotalHorasExtras = HorasExtras + TotalHorasExtras
'                                    TotalExtras = CDate(Horas) + CDate(TotalExtras)
                                    
                                    
                    
                                '--------------------------------------------------------------------------------------------------------------------------------
                                '--------------------------------REPORTE DE LLEGADAS TARDE ----------------------------------------------------------------------
                                '-----------------------------------------------------------------------------------------------------------------------------
                                   HoraTarde = "00:00"
                                   If InTime <> "?" Then
                                        If CDate(HoraEntrada) > (CDate(InTime) + CDate(MinutosTarde)) Then
                                             If ToleranciaTarde = True Then
                                               If InTime <> "?" Then
                                                HoraHorario = CDate(InTime) + CDate("00:00:00")
                                                HorasTarde = DateDiff("S", HoraHorario, HoraEntrada)
                                                HoraTarde = Int(HorasTarde / 3600) & ":" & Int((HorasTarde Mod 3600) / 60)
                                               End If
                                             Else
                                                HoraHorario = CDate(InTime) + CDate(MinutosTarde)
                                                HorasTarde = DateDiff("S", HoraHorario, HoraEntrada)
                                                HoraTarde = Int(HorasTarde / 3600) & ":" & Int((HorasTarde Mod 3600) / 60)
                                             End If
                                         Else
                                            HoraTarde = "00:00"
                                         End If
                                    End If
        
        '                            TotalHorasTarde = CDate(HoraTarde) + CDate(TotalHorasTarde)
                                    
                    
                                    If Me.ChkRestarHorario.Value = 1 Then
                                      If Format(HorasTrabajadas, "##,##0.00") >= HorasExtras Then
                                        HorasLabReal = (CDbl(Format(HorasTrabajadas, "##,##0.00")) - HorasExtras) * 3600
                                        HorasLabReal = ConfHorasTrabajadas * 3600
                                        HoraLaboradas = Int(HorasLabReal / 3600) & ":" & Int((HorasLabReal Mod 3600) / 60)
                                      End If
                                    
                                    End If
                                    
                    
                                     '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                     '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                                      '/////////////////////////////////////////////////////////////////////////////////////
                                    Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                     Me.AdoConsulta.Refresh
                                     If Not Me.AdoConsulta.Recordset.EOF Then
                                    
                                                Select Case j
                                                
                                                    Case 0
                                                    
                                                        Me.AdoReportes.Recordset.AddNew
                                                         Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                                         Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                                         Me.AdoReportes.Recordset("Campo3") = departamento
                                                         Me.AdoReportes.Recordset("Campo15") = HoraLaboradas
                                                         Me.AdoReportes.Recordset("Campo22") = Horas
                                                         Me.AdoReportes.Recordset("Campo7") = HoraTarde
                                                         Me.AdoReportes.Recordset.Update
                                                        Me.AdoReportes.Refresh
                                                     Case 1
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo16") = HoraLaboradas
                                                            Me.AdoBuscaReporte.Recordset("Campo23") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo8") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                         
                                                     Case 2
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo17") = HoraLaboradas
                                                            Me.AdoBuscaReporte.Recordset("Campo24") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo9") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 3
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo18") = HoraLaboradas
                                                            Me.AdoBuscaReporte.Recordset("Campo25") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo10") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 4
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo19") = HoraLaboradas
                                                            Me.AdoBuscaReporte.Recordset("Campo26") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo11") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 5
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo20") = HoraLaboradas
                                                            Me.AdoBuscaReporte.Recordset("Campo27") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo12") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 6
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo21") = HoraLaboradas
                                                            Me.AdoBuscaReporte.Recordset("Campo28") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo5") = Format(TotalTrabajadas, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset("Campo6") = Format(TotalExtras, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset("Campo13") = Format(HoraTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                End Select
                                     End If
                        Next
                        Me.osProgress2.Value = j + 1
                
                   Next
                i = i + 1
                Me.osProgress1.Value = i
                Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
                Me.AdoEmpleados.Recordset.MoveNext
                Me.AdoBuscaReporte.Refresh
              Loop
              
                 
              
        
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.Campo15 AS Dia1, Reportes.Campo16 AS Dia2, Reportes.Campo17 AS Dia3, Reportes.Campo18 AS Dia4, Reportes.Campo19 AS Dia5, Reportes.Campo20 AS Dia6, Reportes.Campo21 AS Dia7, Reportes.Campo22 AS Dia1HE, Reportes.Campo23 AS Dia2HE, Reportes.Campo24 AS Dia3HE, Reportes.Campo25 AS Dia4HE, Reportes.Campo26 AS Dia5HE, Reportes.Campo27 AS Dia6HE, Reportes.Campo28 AS Dia7HE, Reportes.CampoNum1 AS TotalHoras, Reportes.Campo5, Reportes.Campo6 AS TotalHE, Reportes.Campo7 AS Dia1T, Reportes.Campo8 AS Dia2T, Reportes.Campo9 AS Dia3T, Reportes.Campo10 AS Dia4T, Reportes.Campo11 AS Dia5T, Reportes.Campo12 AS Dia6T, Reportes.Campo13 AS Dia7T, Reportes.Campo14 AS TotalTarde From Reportes ORDER BY Reportes.Campo3, Reportes.Campo1"
        
        
          
        
                 Set rpt = New ArepLaboradasExtras
                 rpt.DataControl1.ConnectionString = ConexionReloj
                 rpt.DataControl1.Source = sql
                 fPreview.RunReport rpt
                 fPreview.Show 1
                 
              '*********************************************************************************
              '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
              '*********************************************************************************
              
              rs.Open "DELETE FROM [Reportes] ", ConexionReloj
              
 Case "REPORTE LLEGADAS TARDE SIETE DIAS"
              FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
              FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
              
        
              
              '******************************************************************************
              '//////BUSCO LA CONFIGURACION GENERAL /////////////////////////////////////////
              '*****************************************************************************
               MDIPrimero.DtaEmpresa.Refresh
               If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
                 If MDIPrimero.DtaEmpresa.Recordset("RestarToleranciaLlegada") = True Then
                    ToleranciaTarde = True
                 Else
                    ToleranciaTarde = False
                 End If
               End If
              
              '*********************************************************************************
              '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
              '*********************************************************************************
               rs.Open "DELETE FROM [Reportes] ", ConexionReloj
        
        
              '****************************************************************************************************************************
              '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
              '***************************************************************************************************************************
        '      SQL = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
              If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
                sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
              Else
               sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
                     "WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"
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
                TotalHorasExtras = 0
                TotalExtras = 0
                TotalHorasTarde = "00:00"
                CodigoH = ""
                
                
        
                
                b = 1
                
                  Me.osProgress2.Visible = True
                  Me.osProgress2.Max = 6
                  Me.osProgress2.Min = 0
                  Me.osProgress2.Value = 0
                  
                  TotalHorasTrabajadas = 0
                  TotalTrabajadas = "00:00"
                
                  For j = 0 To 6
                  
                         If j = 0 Then
                            '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                            sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                                  "WHERE ((UserShift.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Me.DTFechaFin.Value, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Me.DTPFechaIni.Value, "YYYY-MM-DD") & "')"
                            Me.AdoConsulta.RecordSource = sql
                            Me.AdoConsulta.Refresh
                            If Not Me.AdoConsulta.Recordset.EOF Then
                              FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                              Ciclo = Me.AdoConsulta.Recordset("Cycles")
                              Date1 = CDate(FechaInicioH)
                              Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                              DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                              FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                            Else
                              Date1 = CDate(Me.DTPFechaIni.Value)
                              Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                              DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                           
                            End If
                         Else
        
                                Date1 = CDate(Me.DTPFechaIni)
                                Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                                DiaInicio = DiaHorario(Date1, Date2, Ciclo)
        
                        End If
                        
                        Me.Caption = "Procesando " & Date2 & " Empleado: " & i & " de " & Me.osProgress1.Max
                        DoEvents
                        
               
                        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '////////////BUSCO EL HORARIO PARA ESTE EMPLEADO ////////////////////////////////////////////////////////////////
                        '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                         Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                       "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                         Me.AdoHorarios.Refresh
                      
                      '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                            If Not Me.AdoHorarios.Recordset.EOF Then
                            
                           
                           
                              Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                            "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(Date2, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Date2, "YYYY-MM-DD") & "')) "
                              Me.AdoHorarios.Refresh
                              If Me.AdoHorarios.Recordset.EOF Then
                                '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                                '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                              "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                                Me.AdoHorarios.Refresh
                                
                                  LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                                   
                                   
                                  If LongitudMinutosIn < 1200 Then  'MENOR A 1400MIN 12 HORAS
                                     '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                                      FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 00:00#"
                                      FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                                      
                                      sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                                    
                                      SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                                  Else
                                      FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                                      FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                                     '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                                      sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                                    
                                      SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                      "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                                   End If
                                   SinHorario = True
                                   HoraInTime(CantHorarios) = "?"
                                   HoraOutTime(CantHorarios) = "?"
                                   SqlIN(0) = sql
                                   SqlOut(0) = SQlSalida
                                   CantHorarios = 1
                              Else
                                   TieneJornadas = False
                                   SinHorario = False
                                   CantHorarios = 0
                                   Me.AdoHorarios.Refresh
                                   If CodEmpleado = 2 Then
                                     CodEmpleado = 2
                                   End If
                                
                                   Do While Not Me.AdoHorarios.Recordset.EOF
                                   
                                       BInTime = Me.AdoHorarios.Recordset("BIntime")
                                       EInTime = Me.AdoHorarios.Recordset("EIntime")
                                       InTime = Me.AdoHorarios.Recordset("Intime")
                                       LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                                       
        '                               Me.AdoHorarios.Recordset.MoveLast
                                       
                                       BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                                       EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                                       OutTime = Me.AdoHorarios.Recordset("OutTime")
                                       TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                       
                                       
                                       FechaIn = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                                       FechaOut = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                                       
                                       FechaHInicio = "#" & FechaIn & " " & BInTime & "#"
                '                       FechaHFinal = "#" & FechaOut & " " & EInTime & "#"
                                       MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                       MinutosTarde = MinutosSalida & ":00" & ":00"
                                       FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BInTime) + CDate(MinutosTarde)
                                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                                       
                                       sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                        
                                       FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
                '                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
                                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                       MinutosTarde = MinutosSalida & ":00" & ":00"
                                       FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
                                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                                  
                                   
            
                                       HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & InTime))
                                       FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                       MinutosTarde = MinutosSalida & ":00" & ":00"
                                       FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                       FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                                       
                                       SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                        
                                        '********************************************************************************************
                                        '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                        '********************************************************************************************
                                        
                                        CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                        CodigoH = Me.AdoHorarios.Recordset("Schid")
                                    
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
                               If TardePermintido <= 60 Then
                                  MinutosTarde = "00:" & TardePermintido & ":00"
                               End If
                               MinutosTardeHorario(CantHorarios) = MinutosTarde
                               HoraInTime(CantHorarios) = InTime
                               HoraOutTime(CantHorarios) = OutTime
                               SqlIN(CantHorarios) = sql
                               SqlOut(CantHorarios) = SQlSalida
                               CantHorarios = CantHorarios + 1
                               Me.AdoHorarios.Recordset.MoveNext
                             Loop
        
                            
                          End If
                            
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
                               
                               
                              '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                              Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                              Me.AdoBuscaReporte.Refresh
                              If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                  CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                                  HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                                  RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                                  RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                                  JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                                  
                                 
                                  
                                  sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                                
                                  SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                                  
                                  TieneJornadas = True
                             
                              Else
                              
                                  TieneJornadas = False
                                  sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                                
                                  SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                  "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                              End If
                              
                                   SinHorario = True
                                   HoraInTime(CantHorarios) = InTime
                                   HoraOutTime(CantHorarios) = OutTime
                                   SqlIN(0) = sql
                                   SqlOut(0) = SQlSalida
                                   CantHorarios = 1
                              
                            End If
                         If CodEmpleado = 2 Then
                           CodEmpleado = 2
                         End If
                         For L = 0 To CantHorarios - 1
                                MinutosTarde = MinutosTardeHorario(L)
                                InTime = HoraInTime(L)
                                OutTime = HoraOutTime(L)
                                sql = SqlIN(L)
                                SQlSalida = SqlOut(L)

                                HoraHorario = CDate(InTime)
                                    '*********************************************************************************************
                                    '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                                    '*********************************************************************************************
                                    
                                        Entrada = "00:00"
                                        HoraEntrada = "00:00"
                                        If TieneJornadas = True Then
                                        
                                            Me.AdoConsulta.RecordSource = sql
                                            Me.AdoConsulta.Refresh
                                            If Not Me.AdoConsulta.Recordset.EOF Then
                                              Entrada = Me.AdoConsulta.Recordset("CheckTime")
                                              HoraEntrada = Format(Entrada, "hh:mm:ss")
                                            End If
                                       
                                        Else
                                            Me.AdoConsulta.RecordSource = sql
                                            Me.AdoConsulta.Refresh
                                            If Not Me.AdoConsulta.Recordset.EOF Then
                                              Entrada = Me.AdoConsulta.Recordset("CheckTime")
                                              HoraEntrada = Format(Entrada, "hh:mm:ss")
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
                                    '///////////////CON ESTA CONSULTA BUSCO LA HORA DE SALIDA///////////////////////////////////
                                    '*********************************************************************************************
                                        Salida = "00:00"
                                        If TieneJornadas = True Then
                                           
                                             '///////////////////////////////CON ESTAS FECHAS BUSCO LA HORA DE SALIDA DE LA JORNADA ///////////////////
                                             
                                             
                                             HoraSalida = CDate(Entrada) + CDate(CInt(HorasLaborales) & ":00:00")
                                             FechaHInicio = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") - CDate(RangoHora1 & ":00")), "mm/dd/yyyy hh:mm:ss")
                                             FechaHFinal = Format(CDate(Entrada) + (CDate(CInt(HorasLaborales) & ":00:00") + CDate(RangoHora2 & ":00")), "mm/dd/yyyy hh:mm:ss")
                                             HoraSalida = Format(Date2, "mm/dd/yyyy") & " 23:59:59"
                                             HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                                             If JornadaIntercalada = False Then
                                                If CDate(FechaHFinal) > CDate(HoraSalida) Then
                                                   FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                                                End If
                                             End If
                                       
                                            FechaHInicio = "#" & FechaHInicio & "#"
                                            FechaHFinal = "#" & FechaHFinal & "#"
                                            
                                            SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                                    End If
                                    
                              
                                    
                                    '*********************************************************************************************
                                    '///////////////CALCULO LAS HORAS TRABAJADAS///////////////////////////////////
                                    '*********************************************************************************************
                                    RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                    
                                    If Entrada = "00:00" Then
                                      Salida = "00:00"
                                    End If
                                    
                                        HorasTrabajadas = 0
                                        HoraLaboradas = "00:00"
                                        If Salida <> "00:00" Then
                                         If Entrada <> "00:00" Then
                    '                      HorasTrabajadas = (DateDiff("h", Entrada, Salida))
                                           HoraLaboradas = ConvertirSegundos((DateDiff("s", Entrada, Salida)), DiaInicio)
                                           HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60) - RestarAlmuerzo  '/////RESTO UNA HORA DE ALMUERZO //////
                                           HoraSalida = Format(Salida, "hh:mm:ss")
                                           TotalTrabajadas = HoraLaboradas + TotalTrabajadas
                    
                                         
                                         Else
                                            HorasTrabajadas = 0
                                            HoraLaboradas = "00:00"
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
                    '                            RestarAlmuerzo = RestaAlmuerzo(CodigoH)
                                                If TieneJornadas = True Then
                                                   If CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) > HorasLaborales Then
                                                       HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                                       Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
                                                   End If
                                                Else
                                                    If ConfCalcularHorasTrab = False Then
                                                      If SinHorario = False Then
                                                       HorasExtras = (CDbl(((DateDiff("s", HoraSalidaHorario, HoraSalida)) / 3600))) * 3600
                                                       Horas = Int(HorasExtras / 3600) & ":" & Int((HorasExtras Mod 3600) / 60)
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
                                                 Horas = "00:00"
                                             End If
                                            Else
                                             HorasExtras = 0
                                             Horas = "00:00"
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
                                    
                                    HorasExtras = Format(HorasExtras, "##,##0.00")
                                    TotalHorasExtras = HorasExtras + TotalHorasExtras
'                                    TotalExtras = CDate(Horas) + CDate(TotalExtras)
                                    
                                    
                    
                                '--------------------------------------------------------------------------------------------------------------------------------
                                '--------------------------------REPORTE DE LLEGADAS TARDE ----------------------------------------------------------------------
                                '-----------------------------------------------------------------------------------------------------------------------------
                                   HoraTarde = "00:00"
                                   If CDate(HoraEntrada) > (CDate(InTime) + CDate(MinutosTarde)) Then
                                        If ToleranciaTarde = True Then
                                          If InTime <> "?" Then
                                           HoraHorario = CDate(InTime) + CDate("00:00:00")
                                           HorasTarde = DateDiff("S", HoraHorario, HoraEntrada)
                                           HoraTarde = Int(HorasTarde / 3600) & ":" & Int((HorasTarde Mod 3600) / 60)
                                          End If
                                        Else
                                           HoraHorario = CDate(InTime) + CDate(MinutosTarde)
                                           HorasTarde = DateDiff("S", HoraHorario, HoraEntrada)
                                           HoraTarde = Int(HorasTarde / 3600) & ":" & Int((HorasTarde Mod 3600) / 60)
                                        End If
                                    Else
                                       HoraTarde = "00:00"
                                    End If
        
                                    
                                    
                                    TotalHorasTarde = sumaHoras(HoraTarde, TotalHorasTarde)
                    
                    
                                     '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                     '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                                     '/////////////////////////////////////////////////////////////////////////////////////
                                     Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                     Me.AdoConsulta.Refresh
                                     If Not Me.AdoConsulta.Recordset.EOF Then
                                    
                                                Select Case j
                                                
                                                    Case 0
                                                        Me.AdoReportes.Recordset.AddNew
                                                         Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                                         Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                                         Me.AdoReportes.Recordset("Campo3") = departamento
                                                         Me.AdoReportes.Recordset("Campo15") = HoraLaboradas
                                                         Me.AdoReportes.Recordset("CampoFecha8") = Horas
                                                         Me.AdoReportes.Recordset("Campo7") = HoraTarde
                                                         Me.AdoReportes.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                         Me.AdoReportes.Recordset.Update
                                                        Me.AdoReportes.Refresh
                                                     Case 1
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo16") = HoraLaboradas
'                                                            Me.AdoBuscaReporte.Recordset("CampoFecha9") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo8") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                         
                                                     Case 2
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo17") = HoraLaboradas
'                                                            Me.AdoBuscaReporte.Recordset("CampoFecha10") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo9") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 3
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo18") = HoraLaboradas
'                                                            Me.AdoBuscaReporte.Recordset("CampoFecha11") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo10") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 4
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo19") = HoraLaboradas
'                                                            Me.AdoBuscaReporte.Recordset("CampoFecha12") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo11") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 5
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo20") = HoraLaboradas
'                                                            Me.AdoBuscaReporte.Recordset("CampoFecha13") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo12") = HoraTarde
                                                            Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                     Case 6
                                                         Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                                         Me.AdoBuscaReporte.Refresh
                                                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                                            Me.AdoBuscaReporte.Recordset("Campo21") = HoraLaboradas
'                                                            Me.AdoBuscaReporte.Recordset("CampoFecha14") = Horas
                                                            Me.AdoBuscaReporte.Recordset("Campo5") = Format(TotalTrabajadas, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset("Campo6") = Format(TotalExtras, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset("Campo13") = Format(HoraTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset("Campo14") = Format(TotalHorasTarde, "hh:mm")
                                                            Me.AdoBuscaReporte.Recordset.Update
                                                         End If
                                                End Select
                                     End If
                        Next
                        Me.osProgress2.Value = j + 1
                
                   Next
                i = i + 1
                Me.osProgress1.Value = i
                Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
                Me.AdoEmpleados.Recordset.MoveNext
                Me.AdoBuscaReporte.Refresh
              Loop
              
         
         sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.Campo7 AS Dia1, Reportes.Campo8 AS Dia2, Reportes.Campo9 AS Dia3, Reportes.Campo10 AS Dia4, Reportes.Campo11 AS Dia5, Reportes.Campo12 AS Dia6, Reportes.Campo13 AS Dia7, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7, Reportes.CampoNum8 AS TotalHoras From Reportes WHERE (((Reportes.Campo14)<>'00:00')) ORDER BY Reportes.Campo3, Reportes.Campo1"


         Set rpt = New ArepLLegadasTardeSiete
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj


End Select

Me.Frame6.Visible = False
End Sub

Private Sub CmdVerReporte4_Click()
Dim sql As String, CodDptoIni As String, CodDptoFin As String
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
Dim TotalTrabajadas As String, TotalExtras As Date, HorasIn As String, CodigoHorario As String
Dim Horas As String, HoraTarde As String, EntradaAlmuerzo As String, SalidaAlmuerzo As String, EntradaAlmuerzo1 As String, EntradaAlmuerzo2 As String, SalidaAlmuerzo1 As String, SalidaAlmuerzo2 As String, ExcluirSabado As Boolean
Dim SQlEntradaAlmuerzo As String, SqlSalidaAlmuerzo As String, TineJornadas As Boolean
Dim EntradaA As String, SalidaA As String, HoraAlmuerzo As String, DifHoras As Double, DiaExtra As Double, TipoHorasTrabajada As String
Dim Fecha As Date, RestarAlmuerzo As Double, SinHorario As Boolean, TotalHorasTarde As String, ToleranciaTarde As Boolean
Dim MinutosExtra As Double, MinutosHorasExtra As Double, CantHorarios As Double, SqlIN(6) As String, SqlOut(6) As String, L As Double, HoraInTime(6) As String, HoraOutTime(6) As String, MinutosTardeHorario(6) As String
Dim TotalTarde(6) As String

TieneJornadas = False
Me.osProgress2.Visible = False
CodigoH = ""
Me.AdoDatosEmpresa.Refresh

        If Not IsNull(Me.AdoDatosEmpresa.Recordset("MinutosExtra")) Then
         MinutosExtra = Me.AdoDatosEmpresa.Recordset("MinutosExtra")
        Else
         MinutosExtra = 0
        End If
        
        CantHorarios = 0

      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


Me.Frame6.Visible = True
Me.Frame6.Top = 1080


Select Case Me.CmbReportes.Text

 Case "REPORTE MARCAS INCOMPLETAS"
 
 Dim DiaAusente As Boolean
 
 DiaAusente = False
 
 
     
      
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaEntradaInicio = FechaHInicio
      FechaEntradaFin = FechaHFinal
                        
      FechaSalidaInicio = FechaHInicio
      FechaSalidaFin = FechaHFinal
      

      
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
        
        If Me.TDBCombo1.Text <> "" And Me.DBEmpleado2.Text <> "" Then
          sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.Userid) Between '" & Me.TDBCombo1.Text & "' And  '" & Me.DBEmpleado2.Text & "' ) AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"

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
        
'        If CodEmpleado = 421 Then
'          CodEmpleado = 421
'        End If
        
       

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
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
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
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                           
                           ConfCalcularHorasTrab = True
                           ConfHorasTrabajadas = 0
                           
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
                                    
'                                    '///////////////////////////////////////////////////////////////////////////////////////
'                                    '////////////////BUSCO LA HORAS LABORADAS SEGUN EL HORARIO /////////////////////////////
'                                    '//////////////////////////////////////////////////////////////////////////////////////
'                                    If TipoHorasTrabajada = "Horario" Then
'                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
'                                       ConfCalcularHorasTrab = True
'                                       ConfHorasTrabajadas = (DateDiff("n", Me.AdoHorarios.Recordset("Intime"), Me.AdoHorarios.Recordset("OutTime")) / 60) - RestarAlmuerzo
'                                    End If
            
                                
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
                                   If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                    TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                   Else
                                    TardePermintido = 0
                                   End If
                                   
                                   FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                                   MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                   MinutosTarde = MinutosSalida & ":00" & ":00"
                                   FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                                   FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                                   
                                   FechaEntradaInicio = FechaHInicio
                                   FechaEntradaFin = FechaHFinal
                                  
                    
                                   
                                   sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                         "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                    
                                   
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                   '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            
                                           ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                           LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & EInTime))
                                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                           MinutosTarde = MinutosSalida & ":00" & ":00"
                                           LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                                           FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                           FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                                   
            '                       HorasIn = Int(LongitudMinutosIn / 60) & ":" & Int(LongitudMinutosIn Mod 60)
            
            '                       If (CDate(InTime) + CDate(HorasIn)) > CDate(Fecha) Then
            '                        '////SI LA SALIDA ES PARA EL DIA SIGUIENTE PASO PARA EL DIA SIGUIENTE
            '                        FechaHInicio = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & BOutTime & "#"
            '                        FechaHFinal = "#" & Format(DateAdd("d", 1, Format(FechaInicial, "DD/MM/yyyy")), "mm/dd/yyyy") & " " & EOutTime & "#" '+ CDate(MinutosTarde) 'Me.DTFechaFin.Value
            ''                        FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
            '                       End If
            
                                   FechaSalidaInicio = FechaHInicio
                                   FechaSalidaFin = FechaHFinal
                                   
                                   SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                               "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                       
                                   
                                                                       '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       If ConfHorasTrabajadas < 0 Then
                                         FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                         ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                       End If
                                    End If
                       
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
                        
                        FechaEntradaInicio = FechaHInicio
                        FechaEntradaFin = FechaHFinal
                        
                        FechaSalidaInicio = FechaHInicio
                        FechaSalidaFin = FechaHFinal
                        
                        
                       
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
                          
                         
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
'                                 If JornadaIntercalada = False Then
'                                    If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                       FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                    End If
'                                 End If
                           
                                FechaHInicio = "#" & FechaHInicio & "#"
                                FechaHFinal = "#" & FechaHFinal & "#"
                                
                                SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
'                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - 1) - HorasLaborales) * 3600
                                           HorasExtras = (CDbl(((DateDiff("s", Entrada, Salida)) / 3600) - RestarAlmuerzo) - HorasLaborales) * 3600
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
                  If CodEmpleado = "610" Then
                   CodEmpleado = "610"
                  End If
               
                  If Entrada = "00:00" And Salida = "00:00" Then
                       DiaAusente = True
                  Else
                       DiaAusente = False
                  End If
                
                            
            If DiaAusente = False Then
               If TieneJustificacion(CodEmpleado, FechaEntradaInicio, FechaSalidaFin) = False Then
                   If Entrada = "00:00" Or Salida = "00:00" Then
                        If Me.ChkAcumulado.Value = 0 Then
                                 Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                 Me.AdoConsulta.Refresh
                                 If Not Me.AdoConsulta.Recordset.EOF Then
                                 
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
                                          Me.AdoReportes.Recordset("Campo6") = FechaEntradaInicio
                                          Me.AdoReportes.Recordset("Campo7") = Format(FechaEntradaFin, "dd/mm/yyyy HH:MM")
                                          Me.AdoReportes.Recordset("Campo8") = Format(FechaSalidaInicio, "dd/mm/yyyy HH:MM")
                                          Me.AdoReportes.Recordset("Campo9") = Format(FechaSalidaFin, "dd/mm/yyyy HH:MM")
                                          
                                         
                                         Me.AdoReportes.Recordset.Update
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
                                          Me.AdoBuscaReporte.Recordset("Campo6") = FechaEntradaInicio
                                          Me.AdoBuscaReporte.Recordset("Campo7") = FechaEntradaFin
                                          Me.AdoBuscaReporte.Recordset("Campo8") = FechaSalidaInicio
                                          Me.AdoBuscaReporte.Recordset("Campo9") = FechaSalidaFin
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
      
         
         
        If Me.ChkTodosDptos.Value = 0 Then
            If Me.DBDptoIni.Text = "" Or Me.DBDptoFin.Text = "" Then
             sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin From Reportes ORDER BY Reportes.Campo6, Reportes.CampoNum1,Reportes.CampoFecha1"
             Set rpt = New ArepMarcasIncompletas   'ArepAsistencia
            Else
             sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin " & _
                   "From Reportes WHERE (((Reportes.Campo3) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Reportes.CampoNum1, Reportes.CampoFecha3, Reportes.Campo6,  Reportes.CampoFecha1"
             Set rpt = New ArepMarcasIncompletas 'ArepAsistenciaDpto
            End If
        Else
             sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin " & _
                   "From Reportes ORDER BY Reportes.Campo3,Reportes.CampoNum1,Reportes.Campo6"
             Set rpt = New ArepMarcasIncompletas  'ArepAsistenciaDpto
        End If
        
        
         Quien = "Marcas Incompletas"
         
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj

  
  
  Case "REPORTE TARJETA DE ASISTENCIA X DIA"
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaEntradaInicio = FechaHInicio
      FechaEntradaFin = FechaHFinal
                        
      FechaSalidaInicio = FechaHInicio
      FechaSalidaFin = FechaHFinal
      
    

      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      If Me.TDBCombo1.Text = "" And Me.DBEmpleado2.Text = "" Then
        If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
            sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
'             sql = "SELECT DISTINCT Checkinout.Userid From Checkinout"
        Else
             sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM Dept INNER JOIN (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) ON Dept.Deptid = Userinfo.Deptid WHERE (((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "'))"
        End If
      Else
        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
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
         
         Entrada = "00:00"
         Salida = "00:00"
         EntradaA = "00:00"
         SalidaA = "00:00"
         HorasTrabajadas = "00:00"
         Horas = "00:00"
         
            '********************************************************************************************
            '///////////////CON ESTA CONSULTA BUSCO LOS DATOS DE CONFIGURACION //////////////////////////
            '********************************************************************************************
                MDIPrimero.DtaEmpresa.Refresh
                If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
'                  FechaInicial = Me.DTPFechaIni.Value
                   FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
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
                
       
      If CodEmpleado = "615" Then
        CodEmpleado = "615"
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
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(FechaInicial, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(FechaInicial, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                          
                          LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                         If LongitudMinutosIn < 1200 Then  'Menor a 1400  12horas
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
                              FechaEntradaInicio = FechaHInicio
                              FechaEntradaFin = FechaHFinal
                              FechaSalidaInicio = FechaHInicio
                              FechaSalidaFin = FechaHFinal
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
                              FechaEntradaInicio = FechaHInicio
                              FechaEntradaFin = FechaHFinal
                              FechaSalidaInicio = FechaHInicio
                              FechaSalidaFin = FechaHFinal
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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

                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                CodigoH = Me.AdoHorarios.Recordset("Schid")
                                 '*******************************************************************************************************************
                                 '*********************************BUSCO EL HORARIO DE ALMUERZO *****************************************************
                                 '*******************************************************************************************************************
                                 Me.AdoHorarioAlmuerzo.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                                    
                                    
                                    Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                                    
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       ConfHorasTrabajadas = (DateDiff("n", Me.AdoHorarios.Recordset("Intime"), Me.AdoHorarios.Recordset("OutTime")) / 60) - RestarAlmuerzo
                                    End If
                    
                                    TieneJornadas = False
                                    
                                      BInTime = Me.AdoHorarios.Recordset("BIntime")
                                      EInTime = Me.AdoHorarios.Recordset("EIntime")
                                      InTime = Me.AdoHorarios.Recordset("Intime")
                                      LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                       
                       
'                                    Me.AdoHorarios.Recordset.MoveLast
                                    
                                    BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                                    EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                                    OutTime = Me.AdoHorarios.Recordset("OutTime")
                                    LongitudMinutosOut = Me.AdoHorarios.Recordset("Longtime")
                                    If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                     TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                    Else
                                     TardePermintido = 0
                                    End If
                                    
                                    FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                                    MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                    MinutosTarde = MinutosSalida & ":00" & ":00"
                                    FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                                    FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
     
                                    FechaEntradaInicio = FechaHInicio
                                    FechaEntradaFin = FechaHFinal
                                   
                                   sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                         "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                    
                                   
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                   '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'                                    HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
'                                    FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                                    MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                                    MinutosTarde = MinutosSalida & ":00" & ":00"
'                                    FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                                    FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                       
                                           ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                           LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & EInTime))
                                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                           MinutosTarde = MinutosSalida & ":00" & ":00"
                                           LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                                          FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                           FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                       
                                           FechaSalidaInicio = FechaHInicio
                                           FechaSalidaFin = FechaHFinal
                                    
                                    SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                          
                                                                         '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
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
                                     SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                     
                                     
                                     FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                                     FechaHFinal = CDate(FechaInicial)
                                     FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                                     SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                 
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
                          
                           FechaEntradaInicio = FechaHInicio
                           FechaEntradaFin = FechaHFinal
                           FechaSalidaInicio = FechaHInicio
                           FechaSalidaFin = FechaHFinal
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                              FechaEntradaInicio = FechaHInicio
                              FechaEntradaFin = FechaHFinal
                              FechaSalidaInicio = FechaHInicio
                              FechaSalidaFin = FechaHFinal
                      
                          TieneJornadas = False
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
                            SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                            
                            
                            FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                            FechaHFinal = CDate(FechaInicial)
                            FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                            SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
                            If Not SQlEntradaAlmuerzo = "" Then
                                
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
                            If Not SqlSalidaAlmuerzo = "" Then
                                
                                Me.AdoConsulta.RecordSource = SqlSalidaAlmuerzo
                                Me.AdoConsulta.Refresh
                                If Not Me.AdoConsulta.Recordset.EOF Then
                                  SalidaA = Me.AdoConsulta.Recordset("CheckTime")
                                End If
                                
                                If ExcluirSabado = True Then
                                  If DiaInicio = 6 Then
                                    EntradaA = "00:00"
                                    SalidaA = "00:00"
                                  End If
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
'                                 HoraSalida = Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59"
                                 HoraSalida = FechaHFinal
                                 HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                 If JornadaIntercalada = False Then
'                                    If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                       FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                    End If
'                                 End If
                           
                                FechaHInicio = "#" & FechaHInicio & "#"
                                FechaHFinal = "#" & FechaHFinal & "#"
                                
                                SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                                   If InTime <> "?" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
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
                                    
                                    HoraAlmuerzo = "00:00"
                                    If Not EntradaA = "00:00" Then
                                     If Not SalidaA = "00:00" Then
            '                            DifHoras = (CDbl(((DateDiff("s", EntradaA, SalidaA)) / 3600) - 1)) * 3600
                                        DifHoras = DateDiff("s", EntradaA, SalidaA)
                                        HoraAlmuerzo = Int(DifHoras / 3600) & ":" & Int((DifHoras Mod 3600) / 60)
                                     End If
                                    End If
         
         

                                    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                                    '/////////////////////////////////////////////////////////////////////////////////////
                                    Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                    FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
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
                                             Me.AdoReportes.Recordset("CampoFecha6") = Format(EntradaA, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("CampoFecha7") = Format(SalidaA, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("CampoFecha8") = Format(HoraAlmuerzo, "hh:mm:ss")
                                             Me.AdoReportes.Recordset("Campo6") = FechaEntradaInicio
                                             Me.AdoReportes.Recordset("Campo7") = FechaEntradaFin
                                             Me.AdoReportes.Recordset("Campo8") = FechaSalidaInicio
                                             Me.AdoReportes.Recordset("Campo9") = FechaSalidaFin
                                            Me.AdoReportes.Recordset.Update
                                    End If
                Next
        Contador = Contador + 1
        FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
        Me.osProgress2.Value = Me.osProgress2.Value + 1
        Loop  '////////CON EL ESTE CICLO RECORRO TODOS LOS DIAS SELECCIONADOS /////////
        
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         Me.AdoReportes.Refresh
      
         
         
   If Me.ChkTodosDptos.Value = 0 Then
         If Me.DBDptoIni.Text = "" Or Me.DBDptoFin.Text = "" Then
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin From Reportes ORDER BY Reportes.Campo3, Reportes.Campo1, Reportes.Campo6"  'Reportes.Campo1,Reportes.CampoFecha3,Reportes.Campo3
                Set rpt = New ArepTarjetaAsistencia
''               rpt.GroupHeader2.Visible = False
'                rpt.GroupFooter2.Visible = False
         Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin  " & _
                      "From Reportes WHERE (((Reportes.Campo3) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Reportes.Campo3, Reportes.Campo1,Reportes.Campo6"
                Set rpt = New ArepTarjetaAsistencia
'                rpt.GroupHeader2.Visible = True
'                rpt.GroupFooter2.Visible = True
'                rpt.GroupFooter1.Visible = False
        End If
    Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin  " & _
                      "From Reportes ORDER BY Reportes.Campo3, Reportes.Campo1, Reportes.Campo6"
                Set rpt = New ArepTarjetaAsistencia
'                rpt.GroupHeader2.Visible = True
'                rpt.GroupFooter2.Visible = True
    End If
         
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
      
      


 Case "REPORTE AUSENCIA X DIA"
      
      
      SinHorario = False
 
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaHInicio = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaHFinal = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      FechaEntradaInicio = FechaHInicio
      FechaEntradaFin = FechaHFinal
                        
      FechaSalidaInicio = FechaHInicio
      FechaSalidaFin = FechaHFinal
      

      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      If Me.TDBCombo1.Text = "" And Me.DBEmpleado2.Text = "" Then
        If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
    '        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
'             sql = "SELECT DISTINCT Checkinout.Userid From Checkinout"
             sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM Dept INNER JOIN (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) ON Dept.Deptid = Userinfo.Deptid "
        Else
             sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM Dept INNER JOIN (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) ON Dept.Deptid = Userinfo.Deptid WHERE (((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "'))"
        End If
      Else
'        sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
         sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM Dept INNER JOIN (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) ON Dept.Deptid = Userinfo.Deptid WHERE (((Checkinout.Userid) Between '" & Me.TDBCombo1.Text & "' And '" & Me.DBEmpleado2.Text & "') AND ((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) ORDER BY Checkinout.Userid"
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
         
         Entrada = "00:00"
         Salida = "00:00"
         EntradaA = "00:00"
         SalidaA = "00:00"
         HorasTrabajadas = "00:00"
         Horas = "00:00"
         
            '********************************************************************************************
            '///////////////CON ESTA CONSULTA BUSCO LOS DATOS DE CONFIGURACION //////////////////////////
            '********************************************************************************************
                MDIPrimero.DtaEmpresa.Refresh
                If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
'                  FechaInicial = Me.DTPFechaIni.Value
                   FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
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
                    
                      SinHorario = True
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(FechaInicial, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(FechaInicial, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                          
                          LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                         If LongitudMinutosIn < 1200 Then  'Menor a 1400  12horas
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(FechaInicial, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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

                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
                                CodigoH = Me.AdoHorarios.Recordset("Schid")
                                 '*******************************************************************************************************************
                                 '*********************************BUSCO EL HORARIO DE ALMUERZO *****************************************************
                                 '*******************************************************************************************************************
                                 Me.AdoHorarioAlmuerzo.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                                    
                                    
                                    Me.AdoBuscaReporte.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodigoHorario & "))"
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
                       
                       
'                                    Me.AdoHorarios.Recordset.MoveLast
                                    
                                    BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                                    EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                                    OutTime = Me.AdoHorarios.Recordset("OutTime")
                                    LongitudMinutosOut = Me.AdoHorarios.Recordset("Longtime")
                                    If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                     TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                                     Else
                                      TardePermintido = 0
                                    End If
                                    
                                    FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & BInTime & "#"  'Me.DtpFechaINI.Value
                                    MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                                    MinutosTarde = MinutosSalida & ":00" & ":00"
                                    FechaHFinal = CDate(FechaInicial & " " & BInTime) + CDate(MinutosTarde)  'Me.DTFechaFin.Value
                                    FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
     
                                    FechaEntradaInicio = FechaHInicio
                                    FechaEntradaFin = FechaHFinal
     
                                   sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                         "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                    
                                   
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                   '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                                   '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

'                                    HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
'                                    FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                                    MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                                    MinutosTarde = MinutosSalida & ":00" & ":00"
'                                    FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                                    FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                                           ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                                           
                                           LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & EInTime))
                                           FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                                           MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                                           MinutosTarde = MinutosSalida & ":00" & ":00"
                                           LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                                           HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                                          FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                                           FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                       
                                           FechaSalidaInicio = FechaHInicio
                                           FechaSalidaFin = FechaHFinal
                       
                                    SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                          
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaFin, 2, 11), "mm/dd/yyyy") & " " & OutTime
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
                                     SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                                     
                                     
                                     FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                                     FechaHFinal = CDate(FechaInicial)
                                     FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                                     SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                 
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
                          
                         
                          
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                          
                          TieneJornadas = True
                     
                      Else
                      
                          TieneJornadas = False
                          sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                          "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                        
                          SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
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
                            SQlEntradaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                            
                            
                            FechaHInicio = "#" & Format(FechaInicial, "mm/dd/yyyy") & " " & SalidaAlmuerzo1 & "#"
                            FechaHFinal = CDate(FechaInicial)
                            FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & SalidaAlmuerzo2 & "#"
                            SqlSalidaAlmuerzo = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                                                 "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                            
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
                            If Not SQlEntradaAlmuerzo = "" Then
                                
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
                            If Not SqlSalidaAlmuerzo = "" Then
                                
                                Me.AdoConsulta.RecordSource = SqlSalidaAlmuerzo
                                Me.AdoConsulta.Refresh
                                If Not Me.AdoConsulta.Recordset.EOF Then
                                  SalidaA = Me.AdoConsulta.Recordset("CheckTime")
                                End If
                                
                                If ExcluirSabado = True Then
                                  If DiaInicio = 6 Then
                                    EntradaA = "00:00"
                                    SalidaA = "00:00"
                                  End If
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
'                                 HoraSalida = Format(FechaInicial, "mm/dd/yyyy") & " 23:59:59"
                                 HoraSalida = FechaHFinal
                                 HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                 If JornadaIntercalada = False Then
'                                    If CDate(FechaHFinal) > CDate(HoraSalida) Then
'                                       FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
'                                    End If
'                                 End If
                           
                                FechaHInicio = "#" & FechaHInicio & "#"
                                FechaHFinal = "#" & FechaHFinal & "#"
                                
                                SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
                                   If InTime <> "?" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
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
                                    
                                    HoraAlmuerzo = "00:00"
                                    If Not EntradaA = "00:00" Then
                                     If Not SalidaA = "00:00" Then
            '                            DifHoras = (CDbl(((DateDiff("s", EntradaA, SalidaA)) / 3600) - 1)) * 3600
                                        DifHoras = DateDiff("s", EntradaA, SalidaA)
                                        HoraAlmuerzo = Int(DifHoras / 3600) & ":" & Int((DifHoras Mod 3600) / 60)
                                     End If
                                    End If
         
                                    '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                                    '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                                    '/////////////////////////////////////////////////////////////////////////////////////
                                    Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                                    Me.AdoConsulta.Refresh
                                    If Not Me.AdoConsulta.Recordset.EOF Then
                                      FechaInicial = DateAdd("d", Contador, Me.DTPFechaIni.Value)
                                        If Entrada = "00:00" And Salida = "00:00" Then
                                          If Me.ChkJustificacion.Value = 1 Then
                                            If SinHorario = False Then
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
                                                 Me.AdoReportes.Recordset("CampoFecha6") = Format(EntradaA, "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("CampoFecha7") = Format(SalidaA, "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("CampoFecha8") = Format(HoraAlmuerzo, "hh:mm:ss")
                                                 
                                                 Me.AdoReportes.Recordset("Campo6") = FechaEntradaInicio
                                                 Me.AdoReportes.Recordset("Campo7") = FechaEntradaFin
                                                 Me.AdoReportes.Recordset("Campo8") = FechaSalidaInicio
                                                 Me.AdoReportes.Recordset("Campo9") = FechaSalidaFin
                                                Me.AdoReportes.Recordset.Update
                                            End If
                                            

                                         ElseIf TieneJustificacion(CodEmpleado, FechaEntradaInicio, FechaEntradaFin) = False Then
                                            If SinHorario = False Then
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
                                                 Me.AdoReportes.Recordset("CampoFecha6") = Format(EntradaA, "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("CampoFecha7") = Format(SalidaA, "hh:mm:ss")
                                                 Me.AdoReportes.Recordset("CampoFecha8") = Format(HoraAlmuerzo, "hh:mm:ss")
                                                 
                                                 Me.AdoReportes.Recordset("Campo6") = FechaEntradaInicio
                                                 Me.AdoReportes.Recordset("Campo7") = FechaEntradaFin
                                                 Me.AdoReportes.Recordset("Campo8") = FechaSalidaInicio
                                                 Me.AdoReportes.Recordset("Campo9") = FechaSalidaFin
                                                Me.AdoReportes.Recordset.Update
                                           End If
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
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         Me.AdoReportes.Refresh
      
         
         
   If Me.ChkTodosDptos.Value = 0 Then
         If Me.DBDptoIni.Text = "" Or Me.DBDptoFin.Text = "" Then
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo,Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin From Reportes ORDER BY Reportes.Campo3, Reportes.CampoNum1,Reportes.CampoFecha3, Reportes.CampoFecha1"
                Set rpt = New ArepAusencia
                
                rpt.GroupHeader1.Visible = True
                rpt.GroupFooter1.Visible = True
         Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo , Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin " & _
                      "From Reportes WHERE (((Reportes.Campo3) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Reportes.CampoNum1, Reportes.Campo3, Reportes.CampoFecha3,  Reportes.CampoFecha1"
                Set rpt = New ArepAusencia
                rpt.GroupHeader1.Visible = False
                rpt.GroupFooter1.Visible = False

        End If
    Else
                sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada, Reportes.CampoFecha2 AS Salida, Reportes.Campo4 AS HorasTrabajadas, Reportes.Campo5 AS HorasExtras, Reportes.CampoFecha3 AS FechaMarca,CampoFecha6 As EntradaA,CampoFecha7 As SalidaA,CampoFecha8 As HoraAlmuerzo, Reportes.Campo6 AS FechaEntradaInicio,Reportes.Campo7 AS FechaEntradaFin,Reportes.Campo8 AS FechaSalidaInicio, Reportes.Campo9 AS FechaSalidaFin  " & _
                      "From Reportes ORDER BY Reportes.Campo3,Reportes.CampoNum1, Reportes.CampoFecha3,  Reportes.CampoFecha1"
                Set rpt = New ArepAusencia
                rpt.GroupHeader1.Visible = False
                rpt.GroupFooter1.Visible = False
    End If
         
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj


End Select

Me.Frame6.Visible = False

End Sub

Private Sub CmdVerReporte5_Click()
Dim sql As String, CodDptoIni As String, CodDptoFin As String
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
Dim TotalTrabajadas As String, TotalExtras As Date, HorasIn As String, CodigoHorario As String
Dim Horas As String, HoraTarde As String, EntradaAlmuerzo As String, SalidaAlmuerzo As String, EntradaAlmuerzo1 As String, EntradaAlmuerzo2 As String, SalidaAlmuerzo1 As String, SalidaAlmuerzo2 As String, ExcluirSabado As Boolean
Dim SQlEntradaAlmuerzo As String, SqlSalidaAlmuerzo As String, TineJornadas As Boolean
Dim EntradaA As String, SalidaA As String, HoraAlmuerzo As String, DifHoras As Double, DiaExtra As Double, TipoHorasTrabajada As String
Dim Fecha As Date, RestarAlmuerzo As Double, SinHorario As Boolean, TotalHorasTarde As String, ToleranciaTarde As Boolean
Dim MinutosExtra As Double, MinutosHorasExtra As Double, CantHorarios As Double, SqlIN(6) As String, SqlOut(6) As String, L As Double, HoraInTime(6) As String, HoraOutTime(6) As String, MinutosTardeHorario(6) As String
Dim TotalTarde(6) As String

TieneJornadas = False
Me.osProgress2.Visible = False
CodigoH = ""
Me.AdoDatosEmpresa.Refresh

        If Not IsNull(Me.AdoDatosEmpresa.Recordset("MinutosExtra")) Then
         MinutosExtra = Me.AdoDatosEmpresa.Recordset("MinutosExtra")
        Else
         MinutosExtra = 0
        End If
        
        CantHorarios = 0

      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


Me.Frame6.Visible = True
Me.Frame6.Top = 1080

Select Case Me.CmbReportes.Text
  Case "REPORTE ASISTENCIA CATORCE DIAS"
        FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      
      If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
        'sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) "
         sql = "SELECT DISTINCT Checkinout.Userid, Userinfo.Name FROM Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) "

      Else
'       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
'             "WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"

       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName, Userinfo.Name FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
             "WHERE (((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) "

      End If
      
      
      If Me.ChkOrdenado.Value = 1 Then
        sql = sql & " ORDER BY Checkinout.Userid"
      Else
        sql = sql & " ORDER BY Userinfo.Name"
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
      
     

      FechaInicial = Me.DTPFechaIni.Value
      
      Do While Not Me.AdoEmpleados.Recordset.EOF
        DoEvents
        
        CodEmpleado = Me.AdoEmpleados.Recordset("Userid")
        
                If CodEmpleado = "168" Then
                  CodEmpleado = "168"
                End If
        b = 1
        
          Me.osProgress2.Visible = True
          Me.osProgress2.Max = 14
          Me.osProgress2.Min = 0
          Me.osProgress2.Value = 0
        
          For j = 0 To 14

                 If j = 0 Then
                    '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                    sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                          "WHERE ((UserShift.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Me.DTFechaFin.Value, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Me.DTPFechaIni.Value, "YYYY-MM-DD") & "')"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                      Ciclo = Me.AdoConsulta.Recordset("Cycles")
                      Date1 = CDate(FechaInicioH)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                    Else
                    
                      Date1 = CDate(Me.DTPFechaIni.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                    End If
                 Else
                      If FechaInicioH = "" Then
                        Date1 = CDate(Me.DTPFechaIni.Value)
                      Else
                        Date1 = CDate(FechaInicioH)
                      End If
'                      Date1 = CDate(Me.DtpFechaINI.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                End If
                
                Me.Caption = "Procesando " & Date2 & " Empleado: " & i & " de " & Me.osProgress1.Max
                DoEvents
                
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
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                  If Not Me.AdoHorarios.Recordset.EOF Then
                    
                    
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(Date2, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Date2, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                        
                        LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                          If LongitudMinutosIn < 720 Then  'MENOR A 1400MIN 12 HORAS
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                        SinHorario = True
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                      Else
                       SinHorario = False
                       TieneJornadas = False
                       CantHorarios = 0
                       Me.AdoHorarios.Refresh
                        
                       Do While Not Me.AdoHorarios.Recordset.EOF
                       
                               BInTime = Me.AdoHorarios.Recordset("BIntime")
                               EInTime = Me.AdoHorarios.Recordset("EIntime")
                               InTime = Me.AdoHorarios.Recordset("Intime")
                               LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                               
'                               Me.AdoHorarios.Recordset.MoveLast
                               
                               BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                               EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                               OutTime = Me.AdoHorarios.Recordset("OutTime")
                               If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                               Else
                                TardePermintido = 0
                               End If
                               
                               
                               FechaIn = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               FechaOut = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               
                               FechaHInicio = "#" & FechaIn & " " & BInTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EInTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BInTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                               
                               sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                               
                               '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                               '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                               '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaIn & " " & InTime))
'                               FechaHInicio = "#" & Mid(CDate(HorasIn), 1, 10) & " " & BOutTime & "#" 'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & Mid(CDate(HorasIn), 1, 10) & " " & EOutTime & "#"
                               
        '                       FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        ''                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
        '                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
        '                       MinutosTarde = MinutosSalida & ":00" & ":00"
        '                       FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
        '                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                               
                               FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                          
                         
    
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & InTime))
'                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                               
                               ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                               LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & EInTime))
                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                              FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                               FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                               
                               SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                           "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                
                            
                                '********************************************************************************************
                                '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                '********************************************************************************************
                                
                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
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
                                
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaInicio, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                    End If
                                
                         
                       SqlIN(CantHorarios) = sql
                       SqlOut(CantHorarios) = SQlSalida
                       CantHorarios = CantHorarios + 1
                       Me.AdoHorarios.Recordset.MoveNext
                     Loop
                    End If
                      
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
                       
                         '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                         Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                         Me.AdoBuscaReporte.Refresh
                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                             CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                             HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                             RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                             RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                             JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                             
                            
                             
                             sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                           
                             SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                             
                             TieneJornadas = True
                        
                         Else
                         
                             TieneJornadas = False
                             sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                           
                             SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                         End If
                         
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                  End If
                

                
                '*********************************************************************************************
                '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                '*********************************************************************************************
               
                For L = 0 To CantHorarios - 1
                        sql = SqlIN(L)
                        SQlSalida = SqlOut(L)
                            
                        If CodEmpleado = "2078" Then
                          CodEmpleado = "2078"
                        End If
                            
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
                            
                           If Entrada <> "00:00" Then
                              If ConfCalcularHorasTrab = True Then
                                  If TipoHorasTrabajada = "HorasTrab" Then
                                    If InTime <> "?" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
                                    End If
                                  End If
                              End If
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
                             HoraSalida = Format(Date2, "mm/dd/yyyy") & " 23:59:59"
                             HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                             If JornadaIntercalada = False Then
                                If CDate(FechaHFinal) > CDate(HoraSalida) Then
                                   FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                                End If
                             End If
                       
                            FechaHInicio = "#" & FechaHInicio & "#"
                            FechaHFinal = "#" & FechaHFinal & "#"
                            
                            SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
    '                HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60)
    '                HoraSalida = Format(Salida, "hh:mm:ss")
    '                If OutTime <> "?" Then
    '                 HoraSalidaHorario = OutTime
    '                 HorasExtras = (DateDiff("n", HoraSalidaHorario, HoraSalida) / 60)
    '                Else
    '                 HorasExtras = 0
    '                End If
    '
    '                If HorasTrabajadas < 0 Then
    '                  HorasTrabajadas = 0
    '                End If
    '
    '                If HorasExtras < 0 Then
    '                  HorasExtras = 0
    '                End If
    
                        If CodEmpleado = "168" Then
                         CodEmpleado = "168"
                        End If
                     
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                        '/////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                    
                            Select Case j
                            
                                Case 0
                                    '-------------------------BUSCO SI EL EMPLEADO YA TIENE REGISTROS --------------------------------
                                    Me.AdoConsultaReporte.RecordSource = "SELECT Reportes.* From Reportes WHERE (((Reportes.Campo1)='" & CodEmpleado & "'))"
                                    Me.AdoConsultaReporte.Refresh
                                    If Me.AdoConsultaReporte.Recordset.EOF Then
                                        Me.AdoReportes.Recordset.AddNew
                                         Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                         Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                         Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                         Me.AdoReportes.Recordset("Campo3") = departamento
                                         Me.AdoReportes.Recordset("CampoFecha1") = Entrada
                                         Me.AdoReportes.Recordset("CampoFecha2") = Salida
                                        Me.AdoReportes.Recordset.Update
                                        Me.AdoReportes.Refresh
                                     End If
                                 Case 1
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha3") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha4") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                     
                                 Case 2
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha5") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha6") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 3
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha7") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha8") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 4
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha9") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha10") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 5
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha11") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha12") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 6
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha13") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha14") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 7
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha15") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha16") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 8
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha17") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha18") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 9
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha19") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha20") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 10
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha21") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha22") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 11
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha23") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha24") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 12
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha25") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha26") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 13
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha27") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha28") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 14
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha29") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha30") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                            End Select
                      End If
                Next
                Me.osProgress2.Value = j + 1
                DoEvents
           Next
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max & "Empleado " & CodEmpleado
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         If Me.ChkOrdenado.Value = 1 Then
          sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada1, Reportes.CampoFecha2 AS Salida1, Reportes.CampoFecha3 AS Entrada2, Reportes.CampoFecha4 AS Salida2, Reportes.CampoFecha5 AS Entrada3, Reportes.CampoFecha6 AS Salida3, Reportes.CampoFecha7 AS Entrada4, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7,Reportes.CampoFecha15 AS Entrada8, Reportes.CampoFecha16 AS Salida8,Reportes.CampoFecha17 AS Entrada9, Reportes.CampoFecha18 AS Salida9, Reportes.CampoFecha19 AS Entrada10, Reportes.CampoFecha20 AS Salida10, Reportes.CampoFecha21 AS Entrada11, Reportes.CampoFecha22 AS Salida11, " & _
                " Reportes.CampoFecha23 AS Entrada12, Reportes.CampoFecha24 AS Salida12,Reportes.CampoFecha25 AS Entrada13, Reportes.CampoFecha26 AS Salida13,Reportes.CampoFecha27 AS Entrada14, Reportes.CampoFecha28 AS Salida14, Reportes.CampoFecha29 AS Entrada15, Reportes.CampoFecha30 AS Salida15 From Reportes ORDER BY Reportes.Campo3,Reportes.CampoNum1,Reportes.CampoFecha1"
         Else
          sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada1, Reportes.CampoFecha2 AS Salida1, Reportes.CampoFecha3 AS Entrada2, Reportes.CampoFecha4 AS Salida2, Reportes.CampoFecha5 AS Entrada3, Reportes.CampoFecha6 AS Salida3, Reportes.CampoFecha7 AS Entrada4, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7 ,Reportes.CampoFecha15 AS Entrada8, Reportes.CampoFecha16 AS Salida8,Reportes.CampoFecha17 AS Entrada9, Reportes.CampoFecha18 AS Salida9, Reportes.CampoFecha19 AS Entrada10, Reportes.CampoFecha20 AS Salida10, Reportes.CampoFecha21 AS Entrada11, Reportes.CampoFecha22 AS Salida11, " & _
                "Reportes.CampoFecha23 AS Entrada12, Reportes.CampoFecha24 AS Salida12,Reportes.CampoFecha25 AS Entrada13, Reportes.CampoFecha26 AS Salida13,Reportes.CampoFecha27 AS Entrada14, Reportes.CampoFecha28 AS Salida14, Reportes.CampoFecha29 AS Entrada15, Reportes.CampoFecha30 AS Salida15 From Reportes ORDER BY Reportes.Campo3,Reportes.Campo2, Reportes.CampoNum1,Reportes.CampoFecha1"
         End If
         
         Set rpt = New ArepAsistenciaCatorce
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
      

  Case "REPORTE ASISTENCIA QUINCE DIAS"
     
      
      FechaIni = "#" & Format(Me.DTPFechaIni.Value, "mm/dd/yyyy") & "#"
      FechaFin = "#" & Format(Me.DTFechaFin.Value, "mm/dd/yyyy") & " 23:59:59#"
      
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
       rs.Open "DELETE FROM [Reportes] ", ConexionReloj


      '****************************************************************************************************************************
      '//////////////////////////////CON ESTA CONSULTA BUSCO TODOS LOS EMPLEADOS QUE MARCARON EN LA FECHA INDICADA ////////////////
      '***************************************************************************************************************************
      
      If Me.DBDptoIni.Text = "" And Me.DBDptoFin.Text = "" Then
        'sql = "SELECT DISTINCT Checkinout.Userid From Checkinout WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) "
         sql = "SELECT DISTINCT Checkinout.Userid, Userinfo.Name FROM Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ")) "

      Else
'       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
'             "WHERE (((Checkinout.CheckTime) Between " & FechaIni & " And " & FechaFin & ") AND ((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) ORDER BY Checkinout.Userid"

       sql = "SELECT DISTINCT Checkinout.Userid, Dept.DeptName, Userinfo.Name FROM (Checkinout INNER JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid  " & _
             "WHERE (((Dept.DeptName) Between '" & Me.DBDptoIni.Text & "' And '" & Me.DBDptoFin.Text & "')) "

      End If
      
      
      If Me.ChkOrdenado.Value = 1 Then
        sql = sql & " ORDER BY Checkinout.Userid"
      Else
        sql = sql & " ORDER BY Userinfo.Name"
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
      
     

      FechaInicial = Me.DTPFechaIni.Value
      
      Do While Not Me.AdoEmpleados.Recordset.EOF
        DoEvents
        
        CodEmpleado = Me.AdoEmpleados.Recordset("Userid")
        
                If CodEmpleado = "168" Then
                  CodEmpleado = "168"
                End If
        b = 1
        
          Me.osProgress2.Visible = True
          Me.osProgress2.Max = 14
          Me.osProgress2.Min = 0
          Me.osProgress2.Value = 0
        
          For j = 0 To 14

                 If j = 0 Then
                    '/////////////////CON ESTA CONSULTA BUSCO LA FECHA DE INICIO DEL HORARIO////////////////
                    sql = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, UserShift.Userid, UserShift.BeginDate, UserShift.EndDate FROM ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) INNER JOIN UserShift ON Schedule.Schid = UserShift.Schid  " & _
                          "WHERE ((UserShift.Userid)='" & CodEmpleado & "') AND (UserShift.BeginDate <= '" & Format(Me.DTFechaFin.Value, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Me.DTPFechaIni.Value, "YYYY-MM-DD") & "')"
                    Me.AdoConsulta.RecordSource = sql
                    Me.AdoConsulta.Refresh
                    If Not Me.AdoConsulta.Recordset.EOF Then
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                      Ciclo = Me.AdoConsulta.Recordset("Cycles")
                      Date1 = CDate(FechaInicioH)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                      FechaInicioH = Me.AdoConsulta.Recordset("BeginDate")
                    Else
                    
                      Date1 = CDate(Me.DTPFechaIni.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                    End If
                 Else
                      If FechaInicioH = "" Then
                        Date1 = CDate(Me.DTPFechaIni.Value)
                      Else
                        Date1 = CDate(FechaInicioH)
                      End If
'                      Date1 = CDate(Me.DtpFechaINI.Value)
                      Date2 = DateAdd("D", j, Me.DTPFechaIni.Value)
                      DiaInicio = DiaHorario(Date1, Date2, Ciclo)
                End If
                
                Me.Caption = "Procesando " & Date2 & " Empleado: " & i & " de " & Me.osProgress1.Max
                DoEvents
                
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
                                               "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                 Me.AdoHorarios.Refresh
              
              '/////////////SI TIENE HORARIO BUSCO LOS REGISTROS DE ENTRADAS PARA UN DIA///////////////
                  If Not Me.AdoHorarios.Recordset.EOF Then
                    
                    
                      Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                    "WHERE (((SchTime.BeginDay)=" & DiaInicio & ") AND ((Userinfo.Userid)='" & CodEmpleado & "') AND(UserShift.BeginDate <= '" & Format(Date2, "YYYY-MM-DD") & "') AND (UserShift.EndDate >= '" & Format(Date2, "YYYY-MM-DD") & "')) "
                      Me.AdoHorarios.Refresh
                      If Me.AdoHorarios.Recordset.EOF Then
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '///////////////////////SI NO SE ENCUENTRA QUIERE DECIR QUE SOLO ES UN DIA /////////////////////////////////////////////////////
                        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoHorarios.RecordSource = "SELECT Schedule.Schid, Schedule.Schname, Schedule.Cycles, Schedule.Units,  SchTime.BeginDay, SchTime.Timeid, TimeTable.Timename, TimeTable.Intime, TimeTable.Outtime, TimeTable.BIntime, TimeTable.EIntime, TimeTable.BOuttime, TimeTable.EOuttime, TimeTable.Latetime, TimeTable.Leavetime, TimeTable.WorkDays, TimeTable.Longtime, TimeTable.MustIn, TimeTable.MustOut, TimeTable.IsFreetime, TimeTable.IsOvertime, Userinfo.Name, Userinfo.Userid, UserShift.BeginDate, UserShift.EndDate FROM Userinfo INNER JOIN (UserShift INNER JOIN ((Schedule INNER JOIN SchTime ON Schedule.Schid = SchTime.Schid) INNER JOIN TimeTable ON SchTime.Timeid = TimeTable.Timeid) ON UserShift.Schid = Schedule.Schid) ON Userinfo.Userid = UserShift.Userid  " & _
                                                      "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserShift.BeginDate)<='" & Format(Date2, "YYYY-MM-DD") & "') AND ((UserShift.EndDate)>='" & Format(Date2, "YYYY-MM-DD") & "'))"
                        Me.AdoHorarios.Refresh
                        
                        LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                           
                           
                          If LongitudMinutosIn < 720 Then  'MENOR A 1400MIN 12 HORAS
                             '///////SI EL HORARIO ES MENOR A 12 HORAS /////////////////////////////////
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 00:00#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                          Else
                              FechaHInicio = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                              FechaHFinal = "#" & Format(Date2, "mm/dd/yyyy") & " 23:59#"
                             '///////SI EL HORARIO ES MAYOR DE 12 HORAS Y NOTIENE HORARIO /////////////////////////////////
                              sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & ent & ")
                            
                              SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                              "WHERE (((Checkinout.Userid)='-100') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") ) ORDER BY Checkinout.CheckTime"  'AND ((Checkinout.CheckType)=" & sal & ")
                           End If
                        SinHorario = True
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                      Else
                       SinHorario = False
                       TieneJornadas = False
                       CantHorarios = 0
                       Me.AdoHorarios.Refresh
                        
                       Do While Not Me.AdoHorarios.Recordset.EOF
                       
                               BInTime = Me.AdoHorarios.Recordset("BIntime")
                               EInTime = Me.AdoHorarios.Recordset("EIntime")
                               InTime = Me.AdoHorarios.Recordset("Intime")
                               LongitudMinutosIn = Me.AdoHorarios.Recordset("Longtime")
                               
'                               Me.AdoHorarios.Recordset.MoveLast
                               
                               BOutTime = Me.AdoHorarios.Recordset("BOuttime")
                               EOutTime = Me.AdoHorarios.Recordset("EOuttime")
                               OutTime = Me.AdoHorarios.Recordset("OutTime")
                               If Not IsNull(Me.AdoHorarios.Recordset("Latetime")) Then
                                TardePermintido = Me.AdoHorarios.Recordset("Latetime")
                               Else
                                TardePermintido = 0
                               End If
                               
                               
                               FechaIn = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               FechaOut = Format(DateAdd("D", j, Me.DTPFechaIni.Value), "mm/dd/yyyy")
                               
                               FechaHInicio = "#" & FechaIn & " " & BInTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EInTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BInTime, EInTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BInTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EInTime & "#"
                               
                               sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                     "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                               
                               '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                               '///////////////////////////////VERIFICO SI LA SALIDA ES PARA EL DIA SIGUIENTE ///////////////////////////////////////////////
                               '/////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaIn & " " & InTime))
'                               FechaHInicio = "#" & Mid(CDate(HorasIn), 1, 10) & " " & BOutTime & "#" 'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & Mid(CDate(HorasIn), 1, 10) & " " & EOutTime & "#"
                               
        '                       FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        ''                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
        '                       MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
        '                       MinutosTarde = MinutosSalida & ":00" & ":00"
        '                       FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
        '                       FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                               
                               FechaHInicio = "#" & FechaIn & " " & BOutTime & "#"
        '                       FechaHFinal = "#" & FechaOut & " " & EOutTime & "#"
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               FechaHFinal = CDate(Format(FechaOut, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde)
                               FechaHFinal = "#" & Format(FechaHFinal, "mm/dd/yyyy") & " " & EOutTime & "#"
                          
                         
    
'                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & InTime))
'                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
'                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
'                               MinutosTarde = MinutosSalida & ":00" & ":00"
'                               FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
'                               FechaHFinal = "#" & CDate(Format(HorasIn, "mm/dd/yyyy")) & " " & EOutTime & "#"
                               
                               ' HorasIn = DateAdd("n", LongitudMinutosIn, CDate(FechaInicial & " " & InTime))
                               LongitudMinutosIn = Abs(DateDiff("n", EInTime, BOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, CDate(Date2 & " " & EInTime))
                               FechaHInicio = "#" & Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime & "#"  'Me.DtpFechaINI.Value
                               MinutosSalida = Abs(DateDiff("h", BOutTime, EOutTime))
                               MinutosTarde = MinutosSalida & ":00" & ":00"
                               LongitudMinutosIn = Abs(DateDiff("n", BOutTime, EOutTime))
                               HorasIn = DateAdd("n", LongitudMinutosIn, Mid(HorasIn, 1, 10) & " " & BOutTime)
'                              FechaHFinal = CDate(Format(HorasIn, "mm/dd/yyyy") & " " & BOutTime) + CDate(MinutosTarde) 'Me.DTFechaFin.Value
                               FechaHFinal = "#" & Format(Mid(HorasIn, 1, 10), "mm/dd/yyyy") & " " & EOutTime & "#"
                               
                               SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
                                           "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & "))"
                
                            
                                '********************************************************************************************
                                '///////////////CON ESTA CONSULTA BUSCO CONFIGURACION HORAS EXTRA//////////////////////////
                                '********************************************************************************************
                                
                                CodigoHorario = Me.AdoHorarios.Recordset("Schid")
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
                                
                                    '///////////////////////////////////////////////////////////////////////////////////////
                                    '////////////////BUSCO LA HORAS  LABORADAS SEGUN EL HORARIO /////////////////////////////
                                    '//////////////////////////////////////////////////////////////////////////////////////
                                    If TipoHorasTrabajada = "Horario" Then
                                       RestarAlmuerzo = RestaAlmuerzo(CodigoH, DiaInicio)
                                       ConfCalcularHorasTrab = True
                                       FechaHInicio = Format(Mid(FechaEntradaInicio, 2, 11), "mm/dd/yyyy") & " " & InTime
                                       FechaHFinal = Format(Mid(FechaSalidaInicio, 2, 11), "mm/dd/yyyy") & " " & OutTime
                                       ConfHorasTrabajadas = (DateDiff("n", FechaHInicio, FechaHFinal) / 60) - RestarAlmuerzo
                                    End If
                                
                         
                       SqlIN(CantHorarios) = sql
                       SqlOut(CantHorarios) = SQlSalida
                       CantHorarios = CantHorarios + 1
                       Me.AdoHorarios.Recordset.MoveNext
                     Loop
                    End If
                      
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
                       
                         '//////////////////////////////BUSCO SI ESTE EMPLEADO TIENE JORNADA LABORAL ASIGNADA ///////////////////////////////////
                         Me.AdoBuscaReporte.RecordSource = "SELECT Jornada.*, AsignacionJornada.UserId, AsignacionJornada.NombreEmpleado FROM Jornada INNER JOIN AsignacionJornada ON Jornada.CodigoJornada = AsignacionJornada.CodigoJornada WHERE (((AsignacionJornada.UserId)='" & CodEmpleado & "'))"
                         Me.AdoBuscaReporte.Refresh
                         If Not Me.AdoBuscaReporte.Recordset.EOF Then
                             CodigoJornada = Me.AdoBuscaReporte.Recordset("CodigoJornada")
                             HorasLaborales = Me.AdoBuscaReporte.Recordset("HorasLaborales")
                             RangoHora1 = Me.AdoBuscaReporte.Recordset("RangoHora1")
                             RangoHora2 = Me.AdoBuscaReporte.Recordset("RangoHora2")
                             JornadaIntercalada = Me.AdoBuscaReporte.Recordset("JornadaIntercalada")
                             
                            
                             
                             sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                           
                             SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ")) ORDER BY Checkinout.CheckTime"
                             
                             TieneJornadas = True
                        
                         Else
                         
                             TieneJornadas = False
                             sql = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Ent & "))"
                           
                             SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType  " & _
                             "WHERE (((Checkinout.Userid)='" & CodEmpleado & "') AND ((Checkinout.CheckTime) Between " & FechaHInicio & " And " & FechaHFinal & ") AND ((Checkinout.CheckType)=" & Sal & "))"
                         End If
                         
                        SqlIN(0) = sql
                        SqlOut(0) = SQlSalida
                        CantHorarios = 1
                  End If
                

                
                '*********************************************************************************************
                '///////////////CON ESTA CONSULTA BUSCO LA HORA DE ENTRADA///////////////////////////////////
                '*********************************************************************************************
               
                For L = 0 To CantHorarios - 1
                        sql = SqlIN(L)
                        SQlSalida = SqlOut(L)
                            
                        If CodEmpleado = "2078" Then
                          CodEmpleado = "2078"
                        End If
                            
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
                            
                           If Entrada <> "00:00" Then
                              If ConfCalcularHorasTrab = True Then
                                  If TipoHorasTrabajada = "HorasTrab" Then
                                    If InTime <> "?" Then
                                     If InTime > Format(Entrada, "hh:mm") Then
                                        Entrada = Mid(Entrada, 1, 10) & " " & InTime & ":00 " & Mid(Entrada, 21, 4)
                                     End If
                                    End If
                                  End If
                              End If
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
                             HoraSalida = Format(Date2, "mm/dd/yyyy") & " 23:59:59"
                             HoraSalida = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                             If JornadaIntercalada = False Then
                                If CDate(FechaHFinal) > CDate(HoraSalida) Then
                                   FechaHFinal = Format(HoraSalida, "mm/dd/yyyy hh:mm:ss")
                                End If
                             End If
                       
                            FechaHInicio = "#" & FechaHInicio & "#"
                            FechaHFinal = "#" & FechaHFinal & "#"
                            
                            SQlSalida = "SELECT Checkinout.Logid, Checkinout.Userid, Checkinout.CheckTime, Checkinout.CheckType, Checkinout.Sensorid, Checkinout.WorkType, Checkinout.AttFlag, Userinfo.Name, Userinfo.Deptid, Userinfo.Duty, Dept.DeptName, FingerClient.Clientid, FingerClient.ClientName, Status.StatusText FROM (((Checkinout LEFT JOIN Userinfo ON Checkinout.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) LEFT JOIN FingerClient ON Checkinout.Sensorid = FingerClient.ClientNumber) LEFT JOIN Status ON Status.Statusid = Checkinout.CheckType " & _
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
    '                HorasTrabajadas = (DateDiff("n", Entrada, Salida) / 60)
    '                HoraSalida = Format(Salida, "hh:mm:ss")
    '                If OutTime <> "?" Then
    '                 HoraSalidaHorario = OutTime
    '                 HorasExtras = (DateDiff("n", HoraSalidaHorario, HoraSalida) / 60)
    '                Else
    '                 HorasExtras = 0
    '                End If
    '
    '                If HorasTrabajadas < 0 Then
    '                  HorasTrabajadas = 0
    '                End If
    '
    '                If HorasExtras < 0 Then
    '                  HorasExtras = 0
    '                End If
    
                        If CodEmpleado = "168" Then
                         CodEmpleado = "168"
                        End If
                     
                        '///////////////////////////////////////////////////////////////////////////////////////////////////////////////
                        '/////////////////////////BUSCO SI EL EMPLEADO EXISTE ///////////////////////////////////////
                        '/////////////////////////////////////////////////////////////////////////////////////
                        Me.AdoConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)='" & CodEmpleado & "'))"
                        Me.AdoConsulta.Refresh
                        If Not Me.AdoConsulta.Recordset.EOF Then
                    
                            Select Case j
                            
                                Case 0
                                  '-------------------------BUSCO SI EL EMPLEADO YA TIENE REGISTROS --------------------------------
                                   Me.AdoConsultaReporte.RecordSource = "SELECT Reportes.* From Reportes WHERE (((Reportes.Campo1)='" & CodEmpleado & "'))"
                                   Me.AdoConsultaReporte.Refresh
                                   If Me.AdoConsultaReporte.Recordset.EOF Then
                                        Me.AdoReportes.Recordset.AddNew
                                         Me.AdoReportes.Recordset("Campo1") = CodEmpleado
                                         Me.AdoReportes.Recordset("CampoNum1") = CodEmpleado
                                         Me.AdoReportes.Recordset("Campo2") = NombreEmpleado
                                         Me.AdoReportes.Recordset("Campo3") = departamento
                                         Me.AdoReportes.Recordset("CampoFecha1") = Entrada
                                         Me.AdoReportes.Recordset("CampoFecha2") = Salida
                                        Me.AdoReportes.Recordset.Update
                                        Me.AdoReportes.Refresh
                                   End If
                                 Case 1
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha3") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha4") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                     
                                 Case 2
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha5") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha6") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 3
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha7") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha8") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 4
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha9") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha10") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 5
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha11") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha12") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 6
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha13") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha14") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 7
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha15") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha16") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 8
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha17") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha18") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 9
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha19") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha20") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 10
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha21") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha22") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 11
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha23") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha24") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 12
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha25") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha26") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 13
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha27") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha28") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                                 Case 14
                                     Me.AdoBuscaReporte.RecordSource = "SELECT Reportes.* From Reportes Where (((Reportes.Campo1) = '" & CodEmpleado & "')) ORDER BY Reportes.Campo1"
                                     Me.AdoBuscaReporte.Refresh
                                     If Not Me.AdoBuscaReporte.Recordset.EOF Then
                                        Me.AdoBuscaReporte.Recordset("CampoFecha29") = Entrada
                                        Me.AdoBuscaReporte.Recordset("CampoFecha30") = Salida
                                        Me.AdoBuscaReporte.Recordset.Update
                                     End If
                            End Select
                      End If
                Next
                Me.osProgress2.Value = j + 1
                DoEvents
           Next
        i = i + 1
        Me.osProgress1.Value = i
        Me.Caption = "Procesando " & i & " de " & Me.osProgress1.Max & "Empleado " & CodEmpleado
        Me.AdoEmpleados.Recordset.MoveNext
      Loop
      
         If Me.ChkOrdenado.Value = 1 Then
          sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada1, Reportes.CampoFecha2 AS Salida1, Reportes.CampoFecha3 AS Entrada2, Reportes.CampoFecha4 AS Salida2, Reportes.CampoFecha5 AS Entrada3, Reportes.CampoFecha6 AS Salida3, Reportes.CampoFecha7 AS Entrada4, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7,Reportes.CampoFecha15 AS Entrada8, Reportes.CampoFecha16 AS Salida8,Reportes.CampoFecha17 AS Entrada9, Reportes.CampoFecha18 AS Salida9, Reportes.CampoFecha19 AS Entrada10, Reportes.CampoFecha20 AS Salida10, Reportes.CampoFecha21 AS Entrada11, Reportes.CampoFecha22 AS Salida11, " & _
                " Reportes.CampoFecha23 AS Entrada12, Reportes.CampoFecha24 AS Salida12,Reportes.CampoFecha25 AS Entrada13, Reportes.CampoFecha26 AS Salida13,Reportes.CampoFecha27 AS Entrada14, Reportes.CampoFecha28 AS Salida14, Reportes.CampoFecha29 AS Entrada15, Reportes.CampoFecha30 AS Salida15 From Reportes ORDER BY Reportes.Campo3,Reportes.CampoNum1,Reportes.CampoFecha1"
         Else
          sql = "SELECT Reportes.Campo1 AS CodEmpleado, Reportes.Campo2 AS NombreEmpleado, Reportes.Campo3 AS Departamento, Reportes.CampoFecha1 AS Entrada1, Reportes.CampoFecha2 AS Salida1, Reportes.CampoFecha3 AS Entrada2, Reportes.CampoFecha4 AS Salida2, Reportes.CampoFecha5 AS Entrada3, Reportes.CampoFecha6 AS Salida3, Reportes.CampoFecha7 AS Entrada4, Reportes.CampoFecha8 AS Salida4, Reportes.CampoFecha9 AS Entrada5, Reportes.CampoFecha10 AS Salida5, Reportes.CampoFecha11 AS Entrada6, Reportes.CampoFecha12 AS Salida6, Reportes.CampoFecha13 AS Entrada7, Reportes.CampoFecha14 AS Salida7 ,Reportes.CampoFecha15 AS Entrada8, Reportes.CampoFecha16 AS Salida8,Reportes.CampoFecha17 AS Entrada9, Reportes.CampoFecha18 AS Salida9, Reportes.CampoFecha19 AS Entrada10, Reportes.CampoFecha20 AS Salida10, Reportes.CampoFecha21 AS Entrada11, Reportes.CampoFecha22 AS Salida11, " & _
                "Reportes.CampoFecha23 AS Entrada12, Reportes.CampoFecha24 AS Salida12,Reportes.CampoFecha25 AS Entrada13, Reportes.CampoFecha26 AS Salida13,Reportes.CampoFecha27 AS Entrada14, Reportes.CampoFecha28 AS Salida14, Reportes.CampoFecha29 AS Entrada15, Reportes.CampoFecha30 AS Salida15 From Reportes ORDER BY Reportes.Campo3,Reportes.Campo2, Reportes.CampoNum1,Reportes.CampoFecha1"
         End If
         
         Set rpt = New ArepAsistenciaQuince
         rpt.DataControl1.ConnectionString = ConexionReloj
         rpt.DataControl1.Source = sql
         fPreview.RunReport rpt
         fPreview.Show 1
         
      '*********************************************************************************
      '/////BORRO TODOS LOS REGISTROS DE REPORTES //////////////////////////////////////
      '*********************************************************************************
      
      rs.Open "DELETE FROM [Reportes] ", ConexionReloj
      
 End Select
End Sub

Private Sub Command1_Click()
'  Quien = "DptoFin"
'  MDIPrimero.MousePointer = 11
'  FrmDepartamentoReportes.Show 1
'  MDIPrimero.MousePointer = 0

  Quien = "Departamento"
  FrmConsultaReloj.Show 1
  Me.DBDptoFin.Text = NombreDepartamento
End Sub

Private Sub Command2_Click()
  Quien = "Codigo"
  FrmConsultaReloj.Show 1
  Me.TDBCombo1.Text = FrmConsultaReloj.Codigo
End Sub

Private Sub Command3_Click()
  Quien = "Codigo"
  FrmConsultaReloj.Show 1
  Me.DBEmpleado2.Text = FrmConsultaReloj.Codigo
End Sub

Private Sub DBDptoIni_Change()
Dim FechaFin As Date


End Sub

Private Sub DtpFechaINI_Change()
Dim FechaFin As Double
Select Case Me.CmbReportes.Text
 Case "REPORTE ASISTENCIA CATORCE DIAS"
       FechaFin = DateAdd("D", 13, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
 Case "REPORTE ASISTENCIA QUINCE DIAS"
       FechaFin = DateAdd("D", 14, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
       
 Case "LISTADO EMPLEADOS"
       FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
   
 Case "REPORTE ASISTENCIA X DIA"
       FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
   
 Case "REPORTE LLEGADAS TARDE"
       FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
   
 Case "REPORTE SALIDA ANTICIPADA"
       FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
  Case "REPORTE HORAS LAB EXTRA SIETE DIAS"
       FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
 Case "REPORTE ASISTENCIA SIETE DIAS"
       FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
 Case "REPORTE HORAS LABORADAS SIETE DIAS"
       FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
 Case "REPORTE HORAS EXTRA SIETE DIAS"
       FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
 Case "REPORTE LLEGADAS TARDE SIETE DIAS"
       FechaFin = DateAdd("D", 6, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
  Case "REPORTE DETALLE ASISTENCIA"
     Me.FrameFecha.Visible = True
     FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
     Me.DTFechaFin.Value = FechaFin
 Case "REPORTE ASISTENCIA Y AUSENCIA X DIA"
       FechaFin = DateAdd("D", 0, Me.DTPFechaIni.Value)
       Me.DTFechaFin.Value = FechaFin
End Select
End Sub

'Option Explicit

'Dim Tape As New clsTape
Private Sub Form_Load()



Me.Timer1.Enabled = True
Me.Timer1.Interval = Tape.Speed

Me.FrameDpto.BackColor = RGB(140, 170, 231)
Me.FrameJustificacion.BackColor = RGB(140, 170, 231)
Me.FrameFecha.BackColor = RGB(140, 170, 231)
Me.FrameEmpleado.BackColor = RGB(140, 170, 231)
Me.ChkAcumulado.BackColor = RGB(140, 170, 231)
Me.ChkTodosDptos.BackColor = RGB(140, 170, 231)
Me.ChkOrdenado.BackColor = RGB(140, 170, 231)
Me.ChkJustificacion.BackColor = RGB(140, 170, 231)
Me.ChkRestarHorario.BackColor = RGB(140, 170, 231)
Me.ChkLink.BackColor = RGB(140, 170, 231)


With Me.AdoJustificacion
   .ConnectionString = ConexionEasy
   .RecordSource = "SELECT LeaveClass.Classid, LeaveClass.Classname FROM LeaveClass "
   .Refresh
End With

With Me.AdoDatosEmpresa
   .ConnectionString = ConexionReloj
   .RecordSource = "SELECT DatosEmpresa.* FROM DatosEmpresa"
   .Refresh
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

With Me.AdoConsultaReporte
  .ConnectionString = ConexionReloj
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

Me.lbltitulo.Caption = Quien

Select Case Quien
   
 Case "Reportes Generales"

 Me.CmbReportes.AddItem ("LISTADO EMPLEADOS")
 Me.CmbReportes.AddItem ("LISTADO HORARIOS")
 Me.CmbReportes.AddItem ("LISTADO DE EQUIPOS")

 Case "Reportes Asistencia"
  Me.CmbReportes.AddItem ("REPORTE ASISTENCIA X DIA")
  Me.CmbReportes.AddItem ("REPORTE LLEGADAS TARDE")
  Me.CmbReportes.AddItem ("REPORTE SALIDA ANTICIPADA")
  Me.CmbReportes.AddItem ("REPORTE ASISTENCIA SIETE DIAS")
  Me.CmbReportes.AddItem ("REPORTE ASISTENCIA QUINCE DIAS")
  Me.CmbReportes.AddItem ("REPORTE ASISTENCIA CATORCE DIAS")
  Me.CmbReportes.AddItem ("REPORTE HORAS LABORADAS SIETE DIAS")
  Me.CmbReportes.AddItem ("REPORTE HORAS LAB EXTRA SIETE DIAS")
  Me.CmbReportes.AddItem ("REPORTE HORAS EXTRA SIETE DIAS")
  Me.CmbReportes.AddItem ("REPORTE LLEGADAS TARDE SIETE DIAS")
  Me.CmbReportes.AddItem ("REPORTE DETALLE ASISTENCIA")
  Me.CmbReportes.AddItem ("REPORTE ASISTENCIA Y AUSENCIA X DIA")
  Me.CmbReportes.AddItem ("REPORTE DE JUSTIFICACION")
  Me.CmbReportes.AddItem ("REPORTE DE ENTRADAS Y SALIDAS")
  Me.CmbReportes.AddItem ("REPORTE TARJETA DE ASISTENCIA X DIA")
  Me.CmbReportes.AddItem ("REPORTE AUSENCIA X DIA")
  Me.CmbReportes.AddItem ("REPORTE MARCAS INCOMPLETAS")
  

 Case "EstadosFinancieros"
  Me.CmbReportes.AddItem ("BALANCE GENERAL")
  Me.CmbReportes.AddItem ("BALANCE ACUMULADO")
  Me.CmbReportes.AddItem ("BALANCE HISTORICO")


 
 End Select
 
 Me.DTPFechaIni.Value = Format(Now, "dd/mm/yyyy")
End Sub


Private Sub SmartButton1_Click()

End Sub

Private Sub Timer1_Timer()
On Error GoTo TipoErrs
Dim intWidth As Integer
Dim intLeft As Integer      'Posición izquierda
Dim objImage As Control     'Control Image
Dim objImage1 As Control
Randomize
'Dim intLeft As Integer      'Posición izquierda
    'Dim objImage As Control     'Control Image
    Randomize   ' Inicializa el generaor de números aleatorios.


    ' Obtiene la anchura de la presentación
    intWidth = picTV.Width
    'Llama al método de la clase Tape
    ' para reproducir la cinta.
    Tape.Animate intWidth
    
    ' Obtiene la propiedad Left a partir de la clase
   intLeft = Tape.Left

If img1.Visible = True Then
        img1.Visible = False
        Set objImage = Img2
    Else
        img1.Visible = True
        Set objImage = img1
    End If
    
 If Lb0.Visible = True Then
   Lb1.Visible = True
   Lb0.Visible = False
   
 ElseIf Lb1.Visible = True Then
    Lb1.Visible = False
    Lb2.Visible = True
 ElseIf Lb2.Visible = True Then
    Lb2.Visible = False
    Lb3.Visible = True
ElseIf Lb3.Visible = True Then
    Lb3.Visible = False
    Lb4.Visible = True
ElseIf Lb4.Visible = True Then
    Lb4.Visible = False
    Lb5.Visible = True
  ElseIf Lb5.Visible = True Then
    Lb5.Visible = False
    Lb6.Visible = True
  ElseIf Lb6.Visible = True Then
    Lb6.Visible = False
    Lb7.Visible = True
  ElseIf Lb7.Visible = True Then
    Lb7.Visible = False
    Lb8.Visible = True
  ElseIf Lb8.Visible = True Then
    Lb8.Visible = False
    Lb9.Visible = True
  ElseIf Lb9.Visible = True Then
    Lb9.Visible = False
    Lb10.Visible = True
  ElseIf Lb10.Visible = True Then
    Lb10.Visible = False
    Lb11.Visible = True
  ElseIf Lb11.Visible = True Then
    Lb11.Visible = False
    Lb12.Visible = True
  ElseIf Lb12.Visible = True Then
    Lb12.Visible = False
    Lb13.Visible = True
  ElseIf Lb13.Visible = True Then
    Lb13.Visible = False
    Lb14.Visible = True
  ElseIf Lb14.Visible = True Then
    Lb14.Visible = False
    Lb15.Visible = True
  ElseIf Lb15.Visible = True Then
    Lb15.Visible = False
    Lb0.Visible = True
    
 End If

' Borra la presentación
    picTV.Cls
    ' Muestra la nueva imagen en la nueva posición
    picTV.PaintPicture objImage.Picture, intLeft, 100, 800, 800
 Exit Sub
TipoErrs:
 MsgBox Err.Description
End Sub

Function ExtraerCodigo(Cadena As String) As String
    Dim Caracter As String
    For i = 1 To Len(Cadena) 'busca algun caracter que no sea numero ni punto
        Caracter = Mid(Cadena, i, 1)
        If Asc(Caracter) < vbKey0 Or Asc(Caracter) > vbKey9 Then
            If Asc(Caracter) <> vbKeyDecimal Then
                Exit For
            End If
        Else
            
        End If
    Next i
    
End Function

