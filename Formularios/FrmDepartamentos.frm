VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Begin VB.Form FrmDepartamentosReloj 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamentos"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   12285
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Top             =   5400
      Width           =   5775
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Height          =   495
         Left            =   4680
         TabIndex        =   6
         Top             =   240
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton CmdMover 
         Caption         =   "Mover"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "Procesar"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton SmartButton7 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   11160
      TabIndex        =   1
      Top             =   5640
      Width           =   975
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   4560
      Top             =   1200
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
            Picture         =   "FrmDepartamentos.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentos.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentos.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentos.frx":0CF6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   1080
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
            Picture         =   "FrmDepartamentos.frx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDepartamentos.frx":24D2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TreeView1 
      Height          =   5175
      Left            =   120
      TabIndex        =   0
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
      Left            =   7920
      Top             =   7440
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
      Left            =   480
      Top             =   7080
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
   Begin TrueOleDBGrid80.TDBGrid TDBGridCuentas 
      Bindings        =   "FrmDepartamentos.frx":407C
      Height          =   5175
      Left            =   5880
      TabIndex        =   2
      Top             =   120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   9128
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
      Splits(0).Caption=   "Empleados"
      Splits(0).DividerColor=   14215660
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
      MultiSelect     =   2
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&H80000009&"
      _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2"
      _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36,.fgcolor=&H80000009&"
      _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&H80000005&"
      _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1,.bgcolor=&HFFFFFF&"
      _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4,.alignment=2,.bgcolor=&HC08080&"
      _StyleDefs(20)  =   ":id=22,.fgcolor=&H0&,.bold=-1,.fontsize=975,.italic=0,.underline=0"
      _StyleDefs(21)  =   ":id=22,.strikethrough=0,.charset=0"
      _StyleDefs(22)  =   ":id=22,.fontname=Viner Hand ITC"
      _StyleDefs(23)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.alignment=2,.fgcolor=&H800000&,.bold=-1"
      _StyleDefs(24)  =   ":id=14,.fontsize=900,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(25)  =   ":id=14,.fontname=Garamond"
      _StyleDefs(26)  =   "Splits(0).FooterStyle:id=15,.parent=3"
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
   Begin MSAdodcLib.Adodc AdoEmpleados 
      Height          =   375
      Left            =   4560
      Top             =   7080
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
End
Attribute VB_Name = "FrmDepartamentosReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Registros As Double, KeyPrincipal As String


Private Sub CmdCancelar_Click()
Me.CmdCancelar.Visible = False
Me.CmdProcesar.Visible = False
Me.CmdMover.Visible = True
End Sub

Private Sub CmdMover_Click()
Dim i As Integer
Dim Posicion As Double

Erase empleados()
Registros = 0
Registros = Me.TDBGridCuentas.SelBookmarks.Count

ReDim empleados(Registros)
 
For i = 0 To Registros - 1
  Posicion = Me.TDBGridCuentas.SelBookmarks.item(i)
  empleados(i) = Me.TDBGridCuentas.Columns(0).CellValue(Posicion)
 
Next



Me.CmdCancelar.Visible = True
Me.CmdProcesar.Visible = True
Me.CmdMover.Visible = False
MsgBox "Seleccione el Nuevo Grupo", vbInformation, "Sistema Contable"
End Sub

Private Sub CmdProcesar_Click()
 Dim CodEmpleado As String, CodDepartamento As String, Texto As String

Me.DtaGrupos.Refresh
Criterio = "KeyGrupo='" & KeyPrincipal & "'"
Me.DtaGrupos.Recordset.Find (Criterio)
If Not Me.DtaGrupos.Recordset.EOF Then
 Texto = Me.DtaGrupos.Recordset("DescripcionGrupo")
Else
  Texto = "Error."
End If

CodDepartamento = BuscaDepartamento(Texto)

For i = 0 To Registros - 1
  CodEmpleado = empleados(i)
  
  Me.DtaConsulta.RecordSource = "SELECT Userinfo.*, Userinfo.Userid From Userinfo WHERE (((Userinfo.Userid)= '" & CodEmpleado & "'))"
  Me.DtaConsulta.Refresh
  If Not Me.DtaConsulta.Recordset.EOF Then
    Me.DtaConsulta.Recordset("Deptid") = CodDepartamento
    Me.DtaConsulta.Recordset.Update
  End If
  
Next

Me.CmdCancelar.Visible = False
Me.CmdProcesar.Visible = False
Me.CmdMover.Visible = True

 Me.AdoEmpleados.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid Where (((Dept.Deptid) = " & CodDepartamento & ")) ORDER BY Userinfo.Userid "
 Me.AdoEmpleados.Refresh
 
Me.TDBGridCuentas.Columns(0).Caption = "Còdigo Empleado"
Me.TDBGridCuentas.Columns(0).Width = 2000
Me.TDBGridCuentas.Columns(1).Caption = "Nombre Empleado"
Me.TDBGridCuentas.Columns(1).Width = 4350

End Sub

Private Sub Form_Load()
Registros = 0

MDIPrimero.Skin1.ApplySkin hWnd


Dim NodX As Node
Dim Relatives As String, RelationsShips As String
Dim LLave As String, Texto As String, Imagen1 As Integer
Dim Imagen2 As Integer

 Me.TDBGridCuentas.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.TDBGridCuentas.OddRowStyle.BackColor = &H80000005
 Me.TDBGridCuentas.AlternatingRowStyle = True

With Me.DtaGrupos
   '.DatabaseName = Ruta
   .ConnectionString = ConexionEasy
End With

With Me.DtaConsulta
   '.DatabaseName = Ruta
   .ConnectionString = ConexionEasy
End With

With Me.AdoEmpleados
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

Private Sub SmartButton7_Click()
Unload Me
End Sub

Private Sub TreeView1_Expand(ByVal Node As MSComctlLib.Node)
  Dim Numero As Integer
  Dim Cadena1 As String, Cadena2 As String
  Dim IdDepartamento As Double
  KeyPadre = ""
  KeyHijo = ""
  KeyNodoUltimo = ""
  KeyPrincipal = Mid(Node.Key, 2, 1)
  IdDepartamento = KeyPrincipal
'  KeyPrincipal = Me.TreeView1.SelectedItem

 Me.AdoEmpleados.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid Where (((Dept.Deptid) = " & IdDepartamento & ")) ORDER BY Userinfo.Userid "
 Me.AdoEmpleados.Refresh
 
Me.TDBGridCuentas.Columns(0).Caption = "Còdigo Empleado"
Me.TDBGridCuentas.Columns(0).Width = 2000
Me.TDBGridCuentas.Columns(1).Caption = "Nombre Empleado"
Me.TDBGridCuentas.Columns(1).Width = 4350


If Len(KeyPrincipal) = 1 Then
    NodoBase = True
Else
    NodoBase = False
    KeyPadre = Node.Parent.Key
End If
End Sub

Private Sub TreeView1_NodeClick(ByVal Node As MSComctlLib.Node)
  Dim Numero As Integer
  Dim Cadena1 As String, Cadena2 As String
  Dim IdDepartamento As Double
  KeyPadre = ""
  KeyHijo = ""
  KeyNodoUltimo = ""
  KeyPrincipal = Mid(Node.Key, 2, 1)
  IdDepartamento = KeyPrincipal
'  KeyPrincipal = Me.TreeView1.SelectedItem

 Me.AdoEmpleados.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid Where (((Dept.Deptid) = " & IdDepartamento & ")) ORDER BY Userinfo.Userid "
 Me.AdoEmpleados.Refresh
 
Me.TDBGridCuentas.Columns(0).Caption = "Còdigo Empleado"
Me.TDBGridCuentas.Columns(0).Width = 2000
Me.TDBGridCuentas.Columns(1).Caption = "Nombre Empleado"
Me.TDBGridCuentas.Columns(1).Width = 4350


If Len(KeyPrincipal) = 1 Then
    NodoBase = True
Else
    NodoBase = False
    KeyPadre = Node.Parent.Key
End If
End Sub
