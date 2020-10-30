VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{562E3E04-2C31-4ECE-83F4-4017EEE51D40}#8.0#0"; "todg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#12.0#0"; "Codejock.Controls.v12.0.0.Demo.ocx"
Begin VB.Form FrmConsultaReloj 
   Caption         =   "Consulta de Registros"
   ClientHeight    =   3600
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   3600
   ScaleWidth      =   10410
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.PushButton PushButton1 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Pegar"
      UseVisualStyle  =   -1  'True
   End
   Begin TrueOleDBGrid80.TDBGrid DbgrProducto 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   4895
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
      Splits(0).DividerColor=   14215660
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   0
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=8,.bold=0,.fontsize=825,.italic=0"
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
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   495
      Left            =   9000
      TabIndex        =   2
      Top             =   3000
      Width           =   1335
      _Version        =   786432
      _ExtentX        =   2355
      _ExtentY        =   873
      _StockProps     =   79
      Caption         =   "Cancelar"
      UseVisualStyle  =   -1  'True
   End
   Begin MSAdodcLib.Adodc DtaProductos 
      Height          =   375
      Left            =   360
      Top             =   5640
      Width           =   3975
      _ExtentX        =   7011
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
      Caption         =   "DtaProductos"
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
Attribute VB_Name = "FrmConsultaReloj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cnx As New ADODB.Connection
Private rs As New ADODB.Recordset, rsConexion As New ADODB.Recordset
Private sql As String
Private modal As Boolean
Private getVal As Boolean
Private Id As Integer

Public Codigo As String, Nombre As String

Private Sub DbgrProducto_FilterChange()
On Error GoTo errTdbg
    'Gets called when an action is performed on the filter bar
    Dim col As TrueOleDBGrid80.Column
    Dim cols As TrueOleDBGrid80.Columns
    
    'On Error GoTo errHandler
    On Error Resume Next
    Set cols = Me.DbgrProducto.Columns
    Dim c As Integer
    
    c = DbgrProducto.col
    DbgrProducto.HoldFields
    sql = rs.Filter
    rs.Filter = getFilter(col, cols)
    
    DbgrProducto.col = c
    DbgrProducto.EditActive = True
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
        Select Case rs.Fields(x).Type
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
 Dim sql As String
 
 MDIPrimero.Skin1.ApplySkin hWnd
 Me.DbgrProducto.EvenRowStyle.BackColor = RGB(216, 228, 248)
 Me.DbgrProducto.OddRowStyle.BackColor = &H80000005
 Me.DbgrProducto.AlternatingRowStyle = True


With Me.DtaProductos
  .ConnectionString = ConexionEasy
End With

If cnx.State = adStateClosed Then
    cnx.ConnectionString = ConexionEasy
    cnx.Open
End If

Select Case Quien
       Case "AsignacionEmpleado"
       
              With Me.DtaProductos
                .ConnectionString = ConexionEasy
              End With
             
             sql = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid ORDER BY Userinfo.Name"
             Me.DtaProductos.RecordSource = sql
             Me.DtaProductos.Refresh
             
            With rs
              .CursorLocation = adUseClient
              .Open sql, ConexionEasy, adOpenDynamic, adLockOptimistic
            End With
             
    
            Me.DbgrProducto.DataSource = rs
             Me.DbgrProducto.Columns(0).Width = 1000
             Me.DbgrProducto.Columns(1).Width = 3500
       Case "Asignacion"
       
             With Me.DtaProductos
               .ConnectionString = ConexionReloj
             End With
             
             sql = "SELECT Jornada.* FROM Jornada"
             Me.DtaProductos.RecordSource = sql
             Me.DtaProductos.Refresh
             
            With rs
              .CursorLocation = adUseClient
              .Open sql, ConexionReloj, adOpenDynamic, adLockOptimistic
            End With
             
    
            Me.DbgrProducto.DataSource = rs
             Me.DbgrProducto.Columns(0).Width = 1000
             Me.DbgrProducto.Columns(1).Width = 3500
       Case "Jornadas"
       
             With Me.DtaProductos
               .ConnectionString = ConexionReloj
             End With
             
             sql = "SELECT Jornada.* FROM Jornada"
             Me.DtaProductos.RecordSource = sql
             Me.DtaProductos.Refresh
             
            With rs
              .CursorLocation = adUseClient
              .Open sql, ConexionReloj, adOpenDynamic, adLockOptimistic
            End With
             
    
            Me.DbgrProducto.DataSource = rs
             Me.DbgrProducto.Columns(0).Width = 1000
             Me.DbgrProducto.Columns(1).Width = 3500

       Case "Tarjeta"
       
              With Me.DtaProductos
                .ConnectionString = ConexionEasy
              End With
              sql = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid ORDER BY Userinfo.Name"
              Me.DtaProductos.RecordSource = sql
              Me.DtaProductos.Refresh
              
              With rs
               .CursorLocation = adUseClient
               .Open sql, cnx, adOpenDynamic, adLockOptimistic
             End With
             
             Me.DbgrProducto.DataSource = rs
              Me.DbgrProducto.Columns(0).Width = 1000
              Me.DbgrProducto.Columns(1).Width = 3500
              Me.DbgrProducto.Columns(2).Width = 2000
              Me.DbgrProducto.Columns(3).Width = 2500
              
       Case "Codigo"
       
              With Me.DtaProductos
                .ConnectionString = ConexionEasy
              End With
              sql = "SELECT Userinfo.Userid, Userinfo.Name, Userinfo.Sex, Dept.DeptName FROM Userinfo INNER JOIN Dept ON Userinfo.Deptid = Dept.Deptid ORDER BY Userinfo.Name"
              Me.DtaProductos.RecordSource = sql
              Me.DtaProductos.Refresh
              
              With rs
               .CursorLocation = adUseClient
               .Open sql, cnx, adOpenDynamic, adLockOptimistic
             End With
             
             Me.DbgrProducto.DataSource = rs
              Me.DbgrProducto.Columns(0).Width = 1000
              Me.DbgrProducto.Columns(1).Width = 3500
              Me.DbgrProducto.Columns(2).Width = 2000
              Me.DbgrProducto.Columns(3).Width = 2500
              
    Case "Departamento"
       
              With Me.DtaProductos
                .ConnectionString = ConexionEasy
              End With
              sql = "SELECT Dept.Deptid, Dept.DeptName FROM Dept"
              Me.DtaProductos.RecordSource = sql
              Me.DtaProductos.Refresh
              
              With rs
               .CursorLocation = adUseClient
               .Open sql, cnx, adOpenDynamic, adLockOptimistic
             End With
             
             Me.DbgrProducto.DataSource = rs
              Me.DbgrProducto.Columns(0).Width = 1000
              Me.DbgrProducto.Columns(1).Width = 3500

            
End Select


End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set cnx = Nothing
Set rs = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
End Sub

Private Sub PushButton1_Click()
 Dim Registros As Double, CodigoEmpleado As String, CodigoJornada As String, Marca As Variant
 Dim Nombre As String

Select Case Quien
        Case "Departamento"
           Codigo = rs("Deptid")
           NombreDepartamento = rs("DeptName")

       Case "Codigo"
           Codigo = rs("Userid")
           Nombre = rs("Name")
       Case "Tarjeta"
           FrmAuxiliar.TDBEmpleados.Text = rs("Userid")
           FrmAuxiliar.LblNombres.Caption = rs("Name")
       Case "Jornadas"
           FrmJornadas.TDBCodigo.Text = rs("CodigoJornada")
           FrmJornadas.TxtNombre.Text = rs("NombreJornada")
           FrmJornadas.SSPinNumCuotas.Text = rs("HorasLaborales")
           FrmJornadas.TxtEntrada1.Text = rs("RangoHora1")
           FrmJornadas.TxtEntrada2.Text = rs("RangoHora2")
   
           If rs("JornadaIntercalada") = True Then
                 FrmJornadas.OptIntercalar.Value = True
           Else
                 FrmJornadas.OptJornadaDia.Value = True
           End If
       Case "Asignacion"
           FrmAsignacion.TDBCodigo.Text = rs("CodigoJornada")
           FrmAsignacion.TxtNombre.Text = rs("NombreJornada")
           FrmAsignacion.SSPinNumCuotas.Text = rs("HorasLaborales")
           
       Case "AsignacionEmpleado"
       
           If FrmAsignacion.TDBCodigo.Text = "" Then
             MsgBox "Seleccione una Jornada", vbCritical, "Zeus Reloj"
             Exit Sub
           End If
           
           CodigoJornada = FrmAsignacion.TDBCodigo.Text
           
          
           '/////////////AGREGO TODOS LOS REGISTROS SELECCIONADOS ///////////////////////////////
'           Registros = Me.DbgrProducto.SelBookmarks.Count
           For Each Marca In Me.DbgrProducto.SelBookmarks
           
            Me.DtaProductos.Recordset.Bookmark = Marca
            CodigoEmpleado = Me.DtaProductos.Recordset("UserId")
            Nombre = Me.DtaProductos.Recordset("Name")
            

            FrmAsignacion.AdoConsulta.RecordSource = "SELECT AsignacionJornada.* From AsignacionJornada WHERE (((AsignacionJornada.UserId)='" & CodigoEmpleado & "') AND ((AsignacionJornada.CodigoJornada)='" & CodigoJornada & "'))"
            FrmAsignacion.AdoConsulta.Refresh
            If FrmAsignacion.AdoConsulta.Recordset.EOF Then
              FrmAsignacion.AdoConsulta.Recordset.AddNew
                FrmAsignacion.AdoConsulta.Recordset("UserId") = CodigoEmpleado
                FrmAsignacion.AdoConsulta.Recordset("CodigoJornada") = CodigoJornada
                FrmAsignacion.AdoConsulta.Recordset("NombreEmpleado") = Nombre
              FrmAsignacion.AdoConsulta.Recordset.Update
            End If
             

           Next
           
End Select

Unload Me
End Sub

Private Sub PushButton2_Click()
Unload Me
End Sub
