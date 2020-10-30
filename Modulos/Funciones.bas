Attribute VB_Name = "FuncionesReloj"
Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst _
    As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 _
    As Long, ByVal un2 As Long) As Long

Function BuscaDepartamento(NombreDepartamento As String) As String
  Dim Sql As String
  
  MDIPrimero.AdoConsultaEasyWay.RecordSource = "SELECT Dept.Deptid, Dept.DeptName From Dept WHERE (((Dept.DeptName)='" & NombreDepartamento & "'))"
  MDIPrimero.AdoConsultaEasyWay.Refresh
  If Not MDIPrimero.AdoConsultaEasyWay.Recordset.EOF Then
    BuscaDepartamento = MDIPrimero.AdoConsultaEasyWay.Recordset("Deptid")
  Else
    BuscaDepartamento = ""
  End If


End Function


Function TieneJustificacion(CodEmpleado As String, FechaIni As String, FechaFin As String) As Boolean
    
   QuienJustifica = "JUSTIFICA"

           If CodEmpleado <> "" Then
           
                FrmReportesReloj.AdoConsulta.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid  " & _
                                                            "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & ")) "

                FrmReportesReloj.AdoConsulta.Refresh
                If Not FrmReportesReloj.AdoConsulta.Recordset.EOF Then
                 TieneJustificacion = True
                 If Not IsNull(FrmReportesReloj.AdoConsulta.Recordset("Classname")) Then
                  QuienJustifica = FrmReportesReloj.AdoConsulta.Recordset("Classname")
                 End If
                Else
                 TieneJustificacion = False
                End If
           Else
                TieneJustificacion = False
           End If



End Function
Function TipoJustificacion(CodEmpleado As String, FechaIni As String, FechaFin As String) As String
    
           If CodEmpleado <> "" Then
           
                FrmReportesReloj.AdoConsulta.RecordSource = "SELECT Userinfo.Userid, Userinfo.Name, Dept.DeptName, Dept.Deptid, UserLeave.BeginTime, UserLeave.EndTime, UserLeave.Whys, LeaveClass.Classname FROM LeaveClass RIGHT JOIN ((UserLeave INNER JOIN Userinfo ON UserLeave.Userid = Userinfo.Userid) LEFT JOIN Dept ON Userinfo.Deptid = Dept.Deptid) ON LeaveClass.Classid = UserLeave.LeaveClassid  " & _
                                                            "WHERE (((Userinfo.Userid)='" & CodEmpleado & "') AND ((UserLeave.BeginTime)<=" & FechaFin & ") AND ((UserLeave.EndTime)>=" & FechaIni & ")) "

                FrmReportesReloj.AdoConsulta.Refresh
                If Not FrmReportesReloj.AdoConsulta.Recordset.EOF Then
                 If Not IsNull(FrmReportesReloj.AdoConsulta.Recordset("Classname")) Then
                 TipoJustificacion = FrmReportesReloj.AdoConsulta.Recordset("Classname")
                 End If
                Else
                 TipoJustificacion = " "
                End If
           Else
                TipoJustificacion = " "
           End If



End Function





Function RestaAlmuerzo(CodHorario As String, Dia As Double) As Double
  Dim HoraInicio As String, HoraFin As String, HoraAlmuerzo As Double, RestarAlmuerzo As Boolean, ExcluirSabados As Boolean
  
   If CodHorario = "" Then
    RestaAlmuerzo = 0
       MDIPrimero.DtaEmpresa.Refresh
       If Not MDIPrimero.DtaEmpresa.Recordset.EOF Then
         RestaAlmuerzo = MDIPrimero.DtaEmpresa.Recordset("RestarAlmuerzo")
       End If
     Exit Function
   End If
  
   MDIPrimero.AdoConsulta.RecordSource = "SELECT Horario.* From Horario WHERE (((Horario.Schid)=" & CodHorario & "))"
   MDIPrimero.AdoConsulta.Refresh
   If Not MDIPrimero.AdoConsulta.Recordset.EOF Then
     If Not IsNull(MDIPrimero.AdoConsulta.Recordset("EntradaAlmuerzo")) Then
       HoraInicio = MDIPrimero.AdoConsulta.Recordset("EntradaAlmuerzo")
     End If
     If Not IsNull(MDIPrimero.AdoConsulta.Recordset("SalidaAlmuerzo")) Then
       HoraFin = MDIPrimero.AdoConsulta.Recordset("SalidaAlmuerzo")
     End If
     
       RestarAlmuerzo = MDIPrimero.AdoConsulta.Recordset("RestarAlmuerzo")
       ExcluirSabados = MDIPrimero.AdoConsulta.Recordset("ExcluirSabado")
       
       If RestarAlmuerzo = True Then
          HoraAlmuerzo = DateDiff("n", HoraInicio, HoraFin) / 60
       Else
          HoraAlmuerzo = 0
       End If
       
       If ExcluirSabados = True Then
        Select Case Dia
           Case 6:  HoraAlmuerzo = 0
           Case 0:  HoraAlmuerzo = 0
        End Select
       End If
       
 
    RestaAlmuerzo = HoraAlmuerzo
    
  
   Else
     RestaAlmuerzo = 0
   End If
   
End Function

Function Buscar_Carpeta(Optional Titulo As String, _
                        Optional Path_Inicial As Variant) As String
  
On Local Error GoTo errFunction
      
    Dim objShell As Object
    Dim objFolder As Object
    Dim o_Carpeta As Object
      
    ' Nuevo objeto Shell.Application
    Set objShell = CreateObject("Shell.Application")
      
    On Error Resume Next
    'Abre el cuadro de diálogo para seleccionar
    Set objFolder = objShell.BrowseForFolder( _
                            0, _
                            Titulo, _
                            0, _
                            Path_Inicial)
      
    ' Devuelve solo el nombre de carpeta
    Set o_Carpeta = objFolder.Self
      
    ' Devuelve la ruta completa seleccionada en el diálogo
    Buscar_Carpeta = o_Carpeta.Path
  
Exit Function
'Error
errFunction:
    MsgBox Err.Description, vbCritical
    Buscar_Carpeta = vbNullString
  
End Function


Public Function sumaHoras(H1 As String, H2 As String) As String
    Dim vh1 As Variant
    Dim vh2 As Variant
    Dim intContador As Integer
    Dim vh3(2) As Long
    Dim H3 As String
    
    'Convertir a arrays
    vh1 = Split(H1, ":")
    vh2 = Split(H2, ":")
    
    'Contemplar tambien los segundos
    For intContador = 0 To 2
    
    'Sumar las horas, minutos, segundos
    If intContador <= UBound(vh1) Then vh3(intContador) = Val(vh1(intContador))
    If intContador <= UBound(vh2) Then vh3(intContador) = vh3(intContador) + Val(vh2(intContador))
    Next intContador
    
    'Descontar las cantidades mayores de 60 en 1 y 2
    vh3(1) = vh3(1) + vh3(2) \ 60
    vh3(2) = vh3(2) Mod 60
    vh3(0) = vh3(0) + vh3(1) \ 60
    vh3(1) = vh3(1) Mod 60
    
    'Constuir la cadena a devolver
    sumaHoras = Format(vh3(0), "00") & ":" & Format(vh3(1), "00") & ":" & Format(vh3(2), "00")

End Function




Public Function DiaHorario(FechaIni As Date, FechaFin As Date, Ciclos As Double) As Double
 Dim i As Double, j As Double, Dias As Double, Diffechas As Double, DiaInicio As Double
 Dim n As Double
 
 Dias = Ciclos * 7

 
 DiaInicio = DiaSemana(Day(FechaIni), Month(FechaIni), Year(FechaIni))
 
 '*****************************************************************************************
 '//////////////////CALCULO EL NUMERO DE DIAS ENTRE LAS DOS FECHAS /////////////////////////
 '******************************************************************************************
 Diffechas = DateDiff("d", FechaIni, FechaFin) + 1
 
 i = DiaInicio
 j = 1
 Do While j <= Diffechas
    If i > (Dias - 1) Then
      i = 0
    End If
    
    n = i
    
    i = i + 1
    j = j + 1
  Loop
 
 DiaHorario = n  '+ (Ciclos - 1) * 7

End Function

Public Function ConvertirSegundos(Segundos As Double, Dia As Double) As String
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String
Dim RestarAlmuerzo As Double

If CodigoH <> "" Then
 RestarAlmuerzo = RestaAlmuerzo(CodigoH, Dia)
Else
 RestarAlmuerzo = 0
 RestarAlmuerzo = RestaAlmuerzo(CodigoH, Dia)
End If

If Segundos > 0 Then
    Horas = Segundos / 3600
    
    If Horas > 0 Then '/////RESTO EL ALMUERZO //////////////////
      Horas = (Segundos / 3600) - RestarAlmuerzo
      Segundos = Horas * 3600
    End If
    
    Horas = Int(Segundos / 3600)
    Minutos = Int((Segundos Mod 3600) / 60)
    Cadena = Horas & ":" & Minutos
Else
    Cadena = "00:00"
End If
ConvertirSegundos = Cadena

End Function
Public Function ConvertirS(Segundos As Double) As String
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String

On Error GoTo TipoErrs

If Segundos > 0 Then
    Horas = Int(Segundos / 3600)
    Minutos = Int((Segundos Mod 3600) / 60)
    
    Cadena = Horas & ":" & Minutos
Else
    Cadena = "00:00"
End If
ConvertirS = Cadena

TipoErrs:
 


End Function


Public Function ConvertirSegundosHoras(Segundos As Double) As Double
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String

If Segundos > 0 Then
    Horas = Int(Segundos / 3600)
   
    If Horas > 0 Then '/////RESTO EL ALMUERZO //////////////////
     Horas = Horas - 1
    End If

End If

ConvertirSegundosHoras = Horas

End Function
Public Function ConvertirSegundosMinutos(Segundos As Double) As Double
Dim Horas As Double
Dim Minutos As Double
Dim Cadena As String

On Error GoTo TipoErrs

If Segundos > 0 Then
    Minutos = Int((Segundos Mod 3600) / 60)
End If

Segundos = 0

ConvertirSegundosMinutos = Minutos

TipoErrs:
' MsgBox Err.Description
 

End Function

Public Function DiaSemana(Dias As Double, Mes As Double, Año As Double) As Double
 Dim A As Double, AñoQ As Double, DosD As Double
 Dim b As Double, c As Integer, D As Double, E As Double, F As Double, R As Double
 

 
 '//////////////////////////////////////////////////////////////////////////
 '////////////////////BUSCAMOS EL SIGLO, DEPENDIENDO DEL ANO ///////////////
 '//////////////////////////////////////////////////////////////////////////
' 1700…1799   1800…1899   1900…1999   2000…2099   2100…2199   2200…2299
'     +5         +3          +1           0          -2          -4

    If Año >= 1700 And Año <= 1799 Then
      A = 5
    ElseIf Año >= 1800 And Año <= 1899 Then
      A = 3
    ElseIf Año >= 1900 And Año <= 1999 Then
      A = 1
    ElseIf Año >= 2000 And Año <= 2099 Then
      A = 0
    ElseIf Año >= 2100 And Año <= 2199 Then
      A = -2
    ElseIf Año >= 2200 And Año <= 2299 Then
      A = -4
    End If
    
 '//////////////////////////////////////////////////////////////////////////////////////
 '//////////////////////////////CALCULO EL CUARTO DEL LOS ULTIMOS DIGITOS DEL ANO ////
 '//////////////////////////////////////////////////////////////////////////////////////
    DosD = Mid(Año, 3, 2)
    AñoQ = Int(DosD / 4)
    b = DosD + AñoQ
 
 '////////////////////////////////////////////////////////////////////////////
 '/////////////////////////CALCULO LOS Años BISIESTOS  ////////////////////////
 '////////////////////////////////////////////////////////////////////////////
 ' Años bisiestos: Éstos son los que cumplen que sus dos últimas cifras forman un múltiplo de 4
 '                 (por ejemplo, 1992 o 2004) excepto los terminados en 00. Entre estos últimos sólo son bisiestos los
 '                 múltiplos de cuatrocientos (por ejemplo 2000). Nuestro tercer coeficiente, C depende de ellos:
 '                 si el año es bisiesto, y el mes es enero o febrero el coeficiente será C = –1. En cualquier otro caso C = 0.
 '                 En nuestro ejemplo, como 2007 no es bisiesto tenemos que C = 0.
 
 '/////BUSCO SI SON MULTIPLOS DE CUATRO //////////////////////////////////////
 

    c = 0
    
   If DosD <> "00" Then
         If Val(DosD / 4) - Int(Val(DosD / 4)) = 0 Then
          '///////SI SON ENTEROS SON MULTIPLOS DE  4
          '///AHORA CONSULTO EL MES CORRESPONDIENTE ///////
          If Mes = 1 Or Mes = 2 Then
           c = 1 - 2
          End If
        End If
   Else
         If Val(Año / 400) - Int(Val(Año / 400)) = 0 Then
          '///////SI SON ENTEROS SON MULTIPLOS DE  400
          '///AHORA CONSULTO EL MES CORRESPONDIENTE ///////
          If Mes = 1 Or Mes = 2 Then
           c = 1 - 2
          End If
        End If
   End If
    
 '/////////////////////////////////////////////////////////////////////////////////
 '//////////////////////CALCULO EL FACTOR PARA EL MES //////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////
' Enero   Feb.    Marzo   Abril   Mayo    Junio   Julio   Agosto  Sept.   Oct.    Nov.    Dic.
'  6       2        2       5       0       3       5       1       4      6       2       4
 Select Case Mes
   Case 1: D = 6
   Case 2: D = 2
   Case 3: D = 2
   Case 4: D = 5
   Case 5: D = 0
   Case 6: D = 3
   Case 7: D = 5
   Case 8: D = 1
   Case 9: D = 4
   Case 10: D = 6
   Case 11: D = 2
   Case 12: D = 4
 End Select
 
 '/////////////////////////////////////////////////////////////////////////////////////////
 '/////////////////////CALCULO EL FACTOR DEL DIA ///////////////////////////////////////////
 '/////////////////////////////////////////////////////////////////////////////////////////
 E = Dias

  '/////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////CORREOMOS EL ALGORITMO PARA VER EL DIA /////////////////////////////
  '//////////////////////////////////////////////////////////////////////////////////////////
'  Lunes   Martes  Miércoles   Jueves  Viernes     Sábado  Domingo
'    1        2        3          4       5           6       0

  F = A + b + c + D + E
  R = F - 7
  
  Do While R > 6
   R = R - 7
  Loop

DiaSemana = R

End Function


Public Function Dia(Fecha As Date) As String
 Dim DiaSemana As Double
 
 DiaSemana = Weekday(Fecha)
 Select Case DiaSemana
    Case 1: Dia = "Domingo"
    Case 2: Dia = "Lunes"
    Case 3: Dia = "Martes"
    Case 4: Dia = "Miercoles"
    Case 5: Dia = "Jueves"
    Case 6: Dia = "Viernes"
    Case 7: Dia = "Sabado"
    
 End Select

End Function


Public Function Inicio_Excel() As Boolean
Dim i As Integer
Dim j As Integer

Set objExcel = New Excel.Application
 
objExcel.Visible = True 'lo hacemos visible
objExcel.SheetsInNewWorkbook = 1 'decimos cuantas hojas queremos en el nuevo documento
objExcel.Workbooks.Add ' añadimos el objeto al workbook

End Function


Public Function Formato_Excel(Num_Campos As Integer, Nombre_Campos() As String) As Boolean

With objExcel.ActiveSheet
        
        'Formato de las celdas de los titulos
        .Range(.Cells(3, 1), .Cells(3, Num_Campos - 1)).Borders.LineStyle = xlContinuous
        .Range(.Cells(3, 1), .Cells(3, 9)).Font.Bold = True
        
    For i = 1 To Num_Campos - 1 Step 1
        .Cells(3, i) = Nombre_Campos(i)
    Next i
        'hasta aki pa colocar los titulos
        
        'a partir de aki ta claro que es pa darle el ancho a las celdas ;-)
       
        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 15
        .Columns("C").ColumnWidth = 15
        .Columns("D").ColumnWidth = 45
        .Columns("E").ColumnWidth = 15
        .Columns("F").ColumnWidth = 15
        .Columns("G").ColumnWidth = 15
        .Columns("H").ColumnWidth = 15
        .Columns("I").ColumnWidth = 20
        
   
    
End With
End Function
