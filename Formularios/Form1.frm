VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "SDK Demo"
   ClientHeight    =   7965
   ClientLeft      =   2940
   ClientTop       =   2085
   ClientWidth     =   13965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   13965
   Begin VB.CommandButton Command71 
      Caption         =   "CKT_GetGPRS/CKT_SetGPRS"
      Height          =   375
      Left            =   8400
      TabIndex        =   83
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command70 
      Caption         =   "CKT_ReadRealtimeClocking OFF/ON"
      Height          =   375
      Left            =   8400
      TabIndex        =   82
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton Command18 
      Caption         =   "CKT_NetDaemon"
      Height          =   375
      Left            =   8400
      TabIndex        =   81
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command69 
      Caption         =   "CKT_GetPictureFileHead CKT_GetPictureFile CKT_DelPictureFile"
      Height          =   855
      Left            =   5640
      TabIndex        =   80
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command68 
      Caption         =   "CKT_SetStateChangeInfo"
      Height          =   375
      Left            =   5640
      TabIndex        =   79
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command67 
      Caption         =   "CKT_GetStateChangeInfo"
      Height          =   375
      Left            =   5640
      TabIndex        =   78
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox Text6 
      Height          =   320
      Left            =   11160
      TabIndex        =   77
      Top             =   6480
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   11160
      TabIndex        =   76
      Text            =   "8888"
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   11160
      TabIndex        =   75
      Text            =   "ModfiyPersonInfoLongName"
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      Height          =   320
      Left            =   11160
      TabIndex        =   74
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command66 
      Caption         =   "CKT_AddMessage "
      Height          =   375
      Left            =   11160
      TabIndex        =   73
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command65 
      Caption         =   "CKT_SetComTimeouts"
      Height          =   375
      Left            =   2880
      TabIndex        =   72
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command56 
      Caption         =   "CKT_SetNetTimeouts"
      Height          =   375
      Left            =   2880
      TabIndex        =   71
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command64 
      Caption         =   "CKT_Demo"
      Height          =   375
      Left            =   8400
      TabIndex        =   70
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command63 
      Caption         =   "CKT_SetSleepTime"
      Height          =   375
      Left            =   2880
      TabIndex        =   69
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command62 
      Caption         =   "CKT_GetClockingRecordEx"
      Height          =   375
      Left            =   8400
      TabIndex        =   68
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command61 
      Caption         =   "CKT_ModifyDeviceSno"
      Height          =   375
      Left            =   2880
      TabIndex        =   67
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command60 
      Caption         =   "CKT_SetDaylightSavingTime"
      Height          =   375
      Left            =   8400
      TabIndex        =   66
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command59 
      Caption         =   "CKT_GetDaylightSavingTime"
      Height          =   375
      Left            =   8400
      TabIndex        =   65
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command58 
      Caption         =   "CKT_ListPersonProgressLongName"
      Height          =   375
      Left            =   11160
      TabIndex        =   64
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command57 
      Caption         =   "CKT_ModifyPersonInfoLongName"
      Height          =   375
      Left            =   11160
      TabIndex        =   63
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command55 
      Caption         =   "CKT_SetDateTimeFormat"
      Height          =   375
      Left            =   8400
      TabIndex        =   62
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command54 
      Caption         =   "CKT_SetKQState"
      Height          =   375
      Left            =   8400
      TabIndex        =   61
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command53 
      Caption         =   "CKT_GetKQState"
      Height          =   375
      Left            =   8400
      TabIndex        =   60
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command52 
      Caption         =   "CKT_GetAllMessageHead"
      Height          =   375
      Left            =   11160
      TabIndex        =   59
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command51 
      Caption         =   "CKT_DelMessageByIndex"
      Height          =   375
      Left            =   11160
      TabIndex        =   58
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton Command50 
      Caption         =   "CKT_SetHitRingInfo"
      Height          =   375
      Left            =   2880
      TabIndex        =   57
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton Command49 
      Caption         =   "CKT_GetMachineNumber"
      Height          =   375
      Left            =   2880
      TabIndex        =   56
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command48 
      Caption         =   "CKT_GetHitRingInfo"
      Height          =   375
      Left            =   2880
      TabIndex        =   55
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command47 
      Caption         =   "CKT_SetWorkCode"
      Height          =   375
      Left            =   5640
      TabIndex        =   54
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command46 
      Caption         =   "CKT_ChangeConnectionMode"
      Height          =   375
      Left            =   5640
      TabIndex        =   53
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command45 
      Caption         =   "CKT_SetGroup"
      Height          =   375
      Left            =   2880
      TabIndex        =   52
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command44 
      Caption         =   "CKT_GetGroup"
      Height          =   375
      Left            =   2880
      TabIndex        =   51
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command43 
      Caption         =   "CKT_EnableLog off"
      Height          =   375
      Left            =   8400
      TabIndex        =   50
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command42 
      Caption         =   "CKT_EnableLog on"
      Height          =   375
      Left            =   10920
      TabIndex        =   49
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command41 
      Caption         =   "CKT_SetTimeSection"
      Height          =   375
      Left            =   2880
      TabIndex        =   48
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command40 
      Caption         =   "CKT_GetTimeSection"
      Height          =   375
      Left            =   2880
      TabIndex        =   47
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command39 
      Caption         =   "CKT_PutFPRawDataLoadFile"
      Height          =   375
      Left            =   5640
      TabIndex        =   46
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command38 
      Caption         =   "CKT_GetFPRawDataSaveFile"
      Height          =   375
      Left            =   5640
      TabIndex        =   45
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command37 
      Caption         =   "CKT_SetDeviceClock"
      Height          =   375
      Left            =   120
      TabIndex        =   44
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command36 
      Caption         =   "CKT_SetAutoUpdate"
      Height          =   375
      Left            =   120
      TabIndex        =   43
      Top             =   6000
      Width           =   2655
   End
   Begin VB.CommandButton Command35 
      Caption         =   "CKT_SetRingAllow"
      Height          =   375
      Left            =   120
      TabIndex        =   42
      Top             =   5520
      Width           =   2655
   End
   Begin VB.CommandButton Command25 
      Caption         =   "CKT_SetDeviceMode"
      Height          =   375
      Left            =   120
      TabIndex        =   41
      Top             =   2160
      Width           =   2655
   End
   Begin VB.CommandButton Command23 
      Caption         =   "CKT_SetRepeatKQ"
      Height          =   375
      Left            =   2880
      TabIndex        =   40
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command13 
      Caption         =   "Get Device List"
      Height          =   375
      Left            =   120
      TabIndex        =   39
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CKT_ForceOpenLock"
      Height          =   375
      Left            =   5640
      TabIndex        =   38
      Top             =   5520
      Width           =   2655
   End
   Begin VB.OptionButton Option3 
      Caption         =   "USB"
      Height          =   255
      Left            =   2040
      TabIndex        =   37
      Top             =   120
      Value           =   -1  'True
      Width           =   615
   End
   Begin VB.CommandButton Command21 
      Caption         =   "CKT_GetMessageByIndex"
      Height          =   375
      Left            =   11160
      TabIndex        =   36
      Top             =   6960
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5400
      List            =   "Form1.frx":001F
      TabIndex        =   35
      Text            =   "Combo1"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command34 
      Caption         =   "End"
      Height          =   375
      Left            =   9720
      TabIndex        =   34
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton Command33 
      Caption         =   "Start"
      Height          =   375
      Left            =   8760
      TabIndex        =   33
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   7200
      TabIndex        =   32
      Text            =   "192.168.0.217"
      Top             =   120
      Width           =   1455
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Net"
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   120
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Com"
      Height          =   255
      Left            =   3600
      TabIndex        =   30
      Top             =   120
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   27
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command32 
      Caption         =   "CKT_SetWG"
      Height          =   375
      Left            =   5640
      TabIndex        =   25
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command31 
      Caption         =   "CKT_GetDeviceInfo"
      Height          =   375
      Left            =   120
      TabIndex        =   24
      Top             =   5040
      Width           =   2655
   End
   Begin VB.CommandButton Command30 
      Caption         =   "CKT_ResetDevice"
      Height          =   375
      Left            =   2880
      TabIndex        =   23
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1560
      Top             =   360
   End
   Begin VB.CommandButton Command29 
      Caption         =   "CKT_SetRealtimeMode"
      Height          =   375
      Left            =   8400
      TabIndex        =   22
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command28 
      Caption         =   "CKT_SetDeviceAdminPassword"
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   6480
      Width           =   2655
   End
   Begin VB.CommandButton Command27 
      Caption         =   "CKT_SetSpeakerVolume"
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   7440
      Width           =   2655
   End
   Begin VB.CommandButton Command26 
      Caption         =   "CKT_SetDoor"
      Height          =   375
      Left            =   120
      TabIndex        =   19
      Top             =   6960
      Width           =   2655
   End
   Begin VB.CommandButton Command24 
      Caption         =   "CKT_GetClockingNewRecordEx"
      Height          =   375
      Left            =   8400
      TabIndex        =   18
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command22 
      Caption         =   "CKT_ClearClockingRecord"
      Height          =   375
      Left            =   2880
      TabIndex        =   17
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command20 
      Caption         =   "CKT_GetCounts"
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command19 
      Caption         =   "CKT_ListPersonProgress"
      Height          =   375
      Left            =   11160
      TabIndex        =   15
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command17 
      Caption         =   "           CKT_GetFPRawData                                      CKT_PutFPRawData"
      Height          =   855
      Left            =   5640
      TabIndex        =   14
      Top             =   1680
      Width           =   2655
   End
   Begin VB.CommandButton Command16 
      Caption         =   "CKT_DeleteAllPersonInfo"
      Height          =   375
      Left            =   11160
      TabIndex        =   13
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command15 
      Caption         =   "CKT_DeletePersonInfo"
      Height          =   375
      Left            =   11160
      TabIndex        =   12
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command14 
      Caption         =   "CKT_ModifyPersonInfo"
      Height          =   375
      Left            =   11160
      TabIndex        =   11
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command12 
      Caption         =   "CKT_PutFPTemplateLoadFile"
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   1200
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      Caption         =   "CKT_GetFPTemplateSaveFile"
      Height          =   375
      Left            =   5640
      TabIndex        =   9
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command10 
      Caption         =   "          CKT_GetFPTemplate                                                             CKT_PutFPTemplate"
      Height          =   855
      Left            =   5640
      TabIndex        =   8
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command9 
      Caption         =   "CKT_SetDeviceDate"
      Height          =   375
      Left            =   2880
      TabIndex        =   7
      Top             =   8400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command8 
      Caption         =   "CKT_GetDeviceClock"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CKT_SetDeviceMAC"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   2655
   End
   Begin VB.CommandButton Command6 
      Caption         =   "CKT_SetDeviceServerIPAddr"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4080
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CKT_SetDeviceGateway"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CKT_SetDeviceMask"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3120
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CKT_SetDeviceIPAddr"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2640
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "CKT_GetDeviceNetInfo"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Width           =   2655
   End
   Begin VB.Label Label4 
      Caption         =   "IP Addr:"
      Height          =   255
      Left            =   6480
      TabIndex        =   29
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "COM Port:"
      Height          =   255
      Left            =   4440
      TabIndex        =   28
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ID:"
      Height          =   255
      Left            =   240
      TabIndex        =   26
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub PCopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (Destination As Any, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub PPCopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
            (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function VarPtrArray Lib "msvbvm50.dll" Alias "VarPtr" _
(Var() As Any) As Long
Private Type SYSTEMTIME
        wYear As Integer
        wMonth As Integer
        wDayOfWeek As Integer
        wDay As Integer
        wHour As Integer
        wMinute As Integer
        wSecond As Integer
        wMilliseconds As Integer
End Type
Private Declare Sub GetLocalTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Dim IDNumber As Long

Public Function DEC_to_HEX(ByVal TDec As Long) As String
     Dim a As String
     Dim Dec As Long
     DEC_to_HEX = ""
     Dec = TDec
     Do While Dec > 0
         a = CStr(Dec Mod 16)
         Select Case a
             Case "10": a = "A"
             Case "11": a = "B"
             Case "12": a = "C"
             Case "13": a = "D"
             Case "14": a = "E"
             Case "15": a = "F"
         End Select
         DEC_to_HEX = a & DEC_to_HEX
         Dec = Dec \ 16
     Loop
End Function

Public Function HEX_to_DEC(ByVal Hex As String) As Long
     Dim i As Long
     Dim B As Long
    
     Hex = UCase(Hex)
     For i = 1 To Len(Hex)
         Select Case Mid(Hex, Len(Hex) - i + 1, 1)
             Case "0": B = B + 16 ^ (i - 1) * 0
             Case "1": B = B + 16 ^ (i - 1) * 1
             Case "2": B = B + 16 ^ (i - 1) * 2
             Case "3": B = B + 16 ^ (i - 1) * 3
             Case "4": B = B + 16 ^ (i - 1) * 4
             Case "5": B = B + 16 ^ (i - 1) * 5
             Case "6": B = B + 16 ^ (i - 1) * 6
             Case "7": B = B + 16 ^ (i - 1) * 7
             Case "8": B = B + 16 ^ (i - 1) * 8
             Case "9": B = B + 16 ^ (i - 1) * 9
             Case "A": B = B + 16 ^ (i - 1) * 10
             Case "B": B = B + 16 ^ (i - 1) * 11
             Case "C": B = B + 16 ^ (i - 1) * 12
             Case "D": B = B + 16 ^ (i - 1) * 13
             Case "E": B = B + 16 ^ (i - 1) * 14
             Case "F": B = B + 16 ^ (i - 1) * 15
         End Select
     Next i
     HEX_to_DEC = B
End Function



Private Sub Command18_Click()
Dim ri As Long
If Option2.Value Then
         ri = CKT_NetDaemon() 'if from net
         If ri = 1 Then  'if from USB
            MsgBox ("CKT_NetDaemon OK")
        End If
    End If
End Sub

Private Sub Command21_Click()

        Dim msg As CKT_MessageInfo
       
        If (CKT_GetMessageByIndex(IDNumber, Text6.Text, msg)) = 1 Then
            
            MsgBox msg.PersonID & Chr(10) & _
            msg.Year1 & Chr(32) & msg.Month1 & Chr(32) & msg.Day1 & Chr(10) & _
            msg.Year2 & Chr(32) & msg.Month2 & Chr(32) & msg.Day2 & Chr(10) & _
            msg.msg
            
        
        End If
End Sub

Private Sub Command33_Click()
Dim ri As Long
    
    IDNumber = CLng(Text1.Text)
    If Option1.Value Then
        Call CKT_RegisterSno(IDNumber, Combo1.ListIndex + 1) 'if from com
    End If
    If Option2.Value Then
         ri = CKT_RegisterNet(IDNumber, Text2.Text) 'if from net
         If ri = 1 Then  'if from USB
            MsgBox ("CKT_RegisterNet OK")
        End If
    End If
    If Option3.Value Then
        If (CKT_RegisterUSB(IDNumber, 0)) = 1 Then  'if from USB
            MsgBox ("CKT_RegisterUSB OK")
        End If
    End If
    

End Sub

Private Sub Command34_Click()
    Call CKT_UnregisterSnoNet(IDNumber)

End Sub

Private Sub Command35_Click()
If (CKT_SetRingAllow(IDNumber, 0)) = 1 Then
        MsgBox ("CKT_SetRingAllow OK")
    Else
        MsgBox ("CKT_SetRingAllow fail")
    End If
End Sub

Private Sub Command36_Click()
If (CKT_SetAutoUpdate(IDNumber, 0)) = 1 Then
        MsgBox ("CKT_SetAutoUpdate OK")
    Else
        MsgBox ("CKT_SetAutoUpdate fail")
    End If
End Sub

Private Sub Command37_Click()
    Dim devclock As DATETIMEINFO
    Dim tim As SYSTEMTIME
    GetLocalTime tim
    devclock.Year = tim.wYear
    devclock.Month = tim.wMonth
    devclock.Day = tim.wDay
    devclock.Hour = tim.wHour
    devclock.Minute = tim.wMinute
    devclock.Second = tim.wSecond
    
If (CKT_SetDeviceClock(IDNumber, devclock)) = 1 Then
        MsgBox ("CKT_SetDeviceClock OK")
    Else
        MsgBox ("CKT_SetDeviceClock fail")
    End If
    
End Sub

Private Sub Command38_Click()
    If CKT_GetFPRawDataSaveFile(IDNumber, 2, 0, "C:\1.anv") = 1 Then
        MsgBox "fingerprint data save in C:\1.anv"
    End If
End Sub

Private Sub Command39_Click()
    If CKT_PutFPRawDataLoadFile(IDNumber, 2, 0, "C:\1.anv") = 1 Then
        MsgBox "fingerprint data get from C:\1.anv£¬and download to time&attendance device"
    End If
End Sub

Private Sub Command40_Click()
Dim TSArray As TimeSectEX
Dim X As Long
X = 1
If CKT_GetTimeSection(IDNumber, X, TSArray) Then
    msg$ = msg$ & "(" & TSArray.z1(0) & TSArray.z1(1) & TSArray.z1(2) & TSArray.z1(3) & ")"
    msg$ = msg$ & "(" & TSArray.z2(0) & TSArray.z2(1) & TSArray.z2(2) & TSArray.z2(3) & ")"
    msg$ = msg$ & "(" & TSArray.z3(0) & TSArray.z3(1) & TSArray.z3(2) & TSArray.z3(3) & ")"
    msg$ = msg$ & "(" & TSArray.z4(0) & TSArray.z4(1) & TSArray.z4(2) & TSArray.z4(3) & ")"
    msg$ = msg$ & "(" & TSArray.z5(0) & TSArray.z5(1) & TSArray.z5(2) & TSArray.z5(3) & ")"
    msg$ = msg$ & "(" & TSArray.z6(0) & TSArray.z6(1) & TSArray.z6(2) & TSArray.z6(3) & ")"
    msg$ = msg$ & "(" & TSArray.z7(0) & TSArray.z7(1) & TSArray.z7(2) & TSArray.z7(3) & ")"
    MsgBox msg$
    MsgBox "CKT_GetTimeSection success."
Else
    MsgBox "CKT_GetTimeSection fail."
End If
End Sub

Private Sub Command41_Click()
Dim TSArray(6) As TimeSect
Dim X As Long
X = 1
For i = 0 To 6
    TSArray(i).bHour = 1
    TSArray(i).bMinute = 1
    TSArray(i).eHour = 1
    TSArray(i).eMinute = 1
Next
If CKT_SetTimeSection(IDNumber, X, TSArray(0)) Then
    MsgBox "CKT_SetTimeSection success."
Else
    MsgBox "CKT_SetTimeSection fail."
End If
End Sub

Private Sub Command42_Click()
CKT_EnableLog (1)
End Sub

Private Sub Command43_Click()
    CKT_EnableLog (0)
End Sub

Private Sub Command44_Click()
Dim GGArray(3) As Long
Dim X As Long
X = 2
If CKT_GetGroup(IDNumber, X, GGArray(0)) Then
    msg$ = msg$ & "(" & GGArray(0) & GGArray(1) & GGArray(2) & GGArray(3) & ")"
    MsgBox msg$
    MsgBox "CKT_GetGroup success."
Else
    MsgBox "CKT_GetGroup fail."
End If
End Sub

Private Sub Command45_Click()
Dim GGArray(3) As Long
Dim X As Long
X = 2
For i = 0 To 3
    GGArray(i) = 1
Next
If CKT_SetGroup(IDNumber, X, GGArray(0)) Then
    MsgBox "CKT_SetGroup success."
Else
    MsgBox "CKT_SetGroup fail."
End If
End Sub

Private Sub Command46_Click()
Dim X As Long
Dim ret As Long
X = 1
If (CKT_ChangeConnectionMode(X)) = 1 Then

        MsgBox ("CKT_ChangeConnectionMode OK")
        ret = CKT_NetDaemonWithPort(5010)
    Else
        MsgBox ("CKT_ChangeConnectionMode fail")
    End If
End Sub

Private Sub Command47_Click()
If (CKT_SetWorkCode(IDNumber, 0)) = 1 Then
        MsgBox ("CKT_SetWorkCode OK")
    Else
        MsgBox ("CKT_SetWorkCode fail")
    End If
End Sub

Private Sub Command48_Click()
Dim RTArray(29) As RingTime
If CKT_GetHitRingInfo(IDNumber, RTArray(0)) Then
    For i = 0 To 29
        msg$ = msg$ & "(" & RTArray(i).Hour & RTArray(i).Minute & RTArray(i).Week & ")"
    Next
    MsgBox msg$
    MsgBox "CKT_GetHitRingInfo success."
Else
    MsgBox "CKT_GetHitRingInfo fail."
End If

End Sub

Private Sub Command49_Click()
Dim CodeId As String
CodeId = "0000000011210001"
If (CKT_GetMachineNumber(IDNumber, CodeId)) = 1 Then
        MsgBox (CodeId)
        MsgBox ("CKT_GetMachineNumber OK")
    Else
        MsgBox ("CKT_GetMachineNumber fail")
    End If
End Sub

Private Sub Command50_Click()
Dim RTArray As RingTime
RTArray.Hour = 1
RTArray.Minute = 1
RTArray.Week = 1
If CKT_SetHitRingInfo(IDNumber, 1, RTArray) Then
    MsgBox "CKT_GetHitRingInfo success."
Else
    MsgBox "CKT_GetHitRingInfo fail."
End If

End Sub

Private Sub Command51_Click()
    If (CKT_DelMessageByIndex(IDNumber, Text6.Text)) = 1 Then
        MsgBox ("CKT_DelMessageByIndex OK!")
    Else
        MsgBox ("CKT_DelMessageByIndex fail")
    End If
End Sub

Private Sub Command52_Click()
Dim MHArray(49) As MessageHead
If CKT_GetAllMessageHead(IDNumber, MHArray(0)) Then
    For i = 0 To 49
        If MHArray(i).PersonID >= 0 Then
        msg$ = msg$ & (i) & "(" & MHArray(i).PersonID & MHArray(i).sYear & MHArray(i).sMon & MHArray(i).sDay & MHArray(i).eYear & MHArray(i).eMon & MHArray(i).eDay & ")"
        End If
    Next
    MsgBox msg$
Else
    MsgBox "CKT_GetAllMessageHead fail."
End If

End Sub

Private Sub Command53_Click()
Dim ckqs As CKT_KQState
If (CKT_GetKQState(IDNumber, ckqs)) = 1 Then
    msg$ = msg$ & ckqs.Num
    For i = 0 To ckqs.Num - 1
        msg$ = msg$ & "("
        For j = 0 To 9
            If ckqs.kqmsg(j, i) <= 48 Then
                GoTo Nj
            End If
            msg$ = msg$ & Chr(ckqs.kqmsg(j, i))
        Next j
Nj:     msg$ = msg$ & ")"
    Next i
    MsgBox msg$
    MsgBox ("CKT_GetKQState OK!")
Else
    MsgBox ("CKT_GetKQState fail")
End If
End Sub

Private Sub Command54_Click()
Dim ckqs As CKT_KQState
ckqs.Num = 4
    For i = 0 To ckqs.Num - 1
        For j = 0 To 9
            ckqs.kqmsg(j, i) = 97 + i
        Next j
    Next i

If (CKT_SetKQState(IDNumber, ckqs)) = 1 Then
    MsgBox ("CKT_GetKQState OK!")
Else
    MsgBox ("CKT_GetKQState fail")
End If

End Sub

Private Sub Command55_Click()
Dim X, y As Long
If (CKT_SetDateTimeFormat(IDNumber, 0, 0)) = 1 Then
        MsgBox ("CKT_SetDateTimeFormat OK")
    Else
        MsgBox ("CKT_SetDateTimeFormat fail")
    End If
End Sub

Private Sub Command56_Click()
If CKT_SetNetTimeouts(10 * 1000) Then
    MsgBox "CKT_SetNetTimeouts success."
Else
    MsgBox "CKT_SetNetTimeouts fail."
End If
End Sub

Private Sub Command57_Click()
    Dim mpiRet As Long
    Dim person As PERSONINFOEX
    
    With person
        .CardNo = 0
        .Name = Text4.Text + Chr(0)
        .Password(0) = &H31
        .Password(1) = &H32
        .Password(2) = &H33
        .Password(3) = &H34
        .Password(4) = 0
        .PersonID = Text5.Text
    End With
    
    mpiRet = CKT_ModifyPersonInfoLongName(IDNumber, person)
    If mpiRet = CKT_RESULT_ADDOK Then
         msg$ = msg$ & "Successfully" & Chr(10) & "PersonID:" & person.PersonID & " Name:" & Trim(person.Name)
        MsgBox msg
    ElseIf mpiRet = CKT_RESULT_CHANGEOK Then
        MsgBox "modeify successfully " + "Name:" + person.Name + person.PersonID
    ElseIf mpiRet = CKT_ERROR_MEMORYFULL Then
        MsgBox "memory full"
    Else
        MsgBox "communication failed"
    End If
End Sub

Private Sub Command58_Click()
    Dim RecordCount, RetCount As Long
    Dim pPersons, pLongRun As Long
    Dim person As PERSONINFOEX
    
    If CKT_ListPersonInfoEx(IDNumber, pLongRun) Then
        Do While True
            ret = CKT_ListPersonProgressLongName(pLongRun, RecordCount, RetCount, pPersons)
            If ret = 0 Then
               Exit Sub
            End If
            
            If (ret <> 0) Then
                Dim ptemp As Long
                ptemp = pPersons
                
                For i = 0 To RetCount - 1
                    Call PCopyMemory(person, pPersons, PERSONINFOSIZEEX)
                    person.Name = Left(person.Name, InStr(person.Name, Chr(0)) - 1)
                    pPersons = pPersons + PERSONINFOSIZEEX
                    msg$ = msg$ & "PersonID:" & person.PersonID & "  Name:" & Trim(person.Name)
                    MsgBox "RetCount:" & i + 1 & Chr(10) & msg$
                    msg$ = ""
                Next
                
                If msg$ <> "" Then
                    MsgBox msg$
                End If
                
                If ptemp <> 0 Then
                    Call CKT_FreeMemory(ptemp)
                End If
            End If
            
            If ret <> 2 Then
                Exit Sub
            End If
        Loop
    End If
End Sub

Private Sub Command59_Click()
    Dim dst(15) As Byte
    'Dim dst As String
    Dim i As Integer
    'dst = "0000000000000000"
    For i = 0 To 15
           dst(i) = 0
        Next
    
    If CKT_GetDaylightSavingTime(IDNumber, dst(0)) Then
        For i = 0 To 15
            msg$ = msg$ & dst(i) & ","
        Next
    Else
        msg$ = "Fail !"
    End If
    
    MsgBox msg$
End Sub

Private Sub Command60_Click()
Dim dst(15) As Byte
dst(0) = 1 ' 0-off 1-on
dst(1) = 1 'Mode: 1-Date 2-Week
 '¿ªÊ¼
dst(2) = 2 'month
dst(3) = 29 'day
dst(4) = 0 'start week
dst(5) = 0 'week
dst(6) = 12 'hour
dst(7) = 25 'minute
dst(8) = 1 'second
 '½áÊø
dst(9) = 2 'month
dst(10) = 29 'day
dst(11) = 0 'end week
dst(12) = 0 'week
dst(13) = 14 'hour
dst(14) = 0 'minute
dst(15) = 0 'second

    
    If CKT_SetDaylightSavingTime(IDNumber, dst(0)) Then
        MsgBox ("CKT_SetDaylightSavingTime OK")
    Else
        MsgBox ("CKT_SetDaylightSavingTime fail")
    End If
End Sub

Private Sub Command61_Click()
If CKT_ModifyDeviceSno(IDNumber, 21) Then
        MsgBox ("CKT_ModifyDeviceSno OK")
    Else
        MsgBox ("CKT_ModifyDeviceSno fail")
    End If
End Sub

Private Sub Command62_Click()
    Dim RecordCount, RetCount As Long
    Dim pClockings, pLongRun As Long
    Dim clocking As CLOCKINGRECORD
    
    If CKT_GetClockingRecordEx(IDNumber, pLongRun) Then
        Do While True
            ret = CKT_GetClockingRecordProgress(pLongRun, RecordCount, RetCount, pClockings)
            If ret = 0 Then
               Exit Sub
            End If
            
            If (ret <> 0) Then
                Dim ptemp As Long
                ptemp = pClockings
                
                For i = 1 To RetCount
                    Call PCopyMemory(clocking, pClockings, CLOCKINGRECORDSIZE)
                    pClockings = pClockings + CLOCKINGRECORDSIZE
                    
                    msg$ = msg$ & clocking.PersonID & "  "
                    'If i Mod 50 = 0 Then
                    '    MsgBox "RetCount:" & RetCount & Chr(10) & msg$
                    '    msg$ = ""
                    'End If
                Next
                
                If msg$ <> "" Then
                    MsgBox msg$
                End If
                
                If ptemp <> 0 Then
                    Call CKT_FreeMemory(ptemp)
                End If
            End If
            
            If ret = 1 Then
                Exit Sub
            End If
        Loop
    End If
End Sub

Private Sub Command63_Click()
If CKT_SetSleepTime(IDNumber, 0) Then '0-255
        MsgBox ("CKT_SetSleepTime OK")
    Else
        MsgBox ("CKT_SetSleepTime fail")
    End If
End Sub

Private Sub Command64_Click()
    Dim RecordCount, RetCount As Long
    Dim pClockings, pLongRun As Long
    Dim clocking As CLOCKINGRECORD
    
    Dim ri As Long
    
    '11
    IDNumber = CLng(Text1.Text)
    If Option1.Value Then
        Call CKT_RegisterSno(IDNumber, Combo1.ListIndex + 1) 'if from com
    End If
    If Option2.Value Then
         ri = CKT_RegisterNet(IDNumber, Text2.Text) 'if from net
         'If ri = 1 Then  'if from USB
         '   MsgBox ("CKT_RegisterNet OK")
        'End If
    End If
    If Option3.Value Then
        If (CKT_RegisterUSB(IDNumber, 0)) = 1 Then  'if from USB
            MsgBox ("CKT_RegisterUSB OK")
        End If
    End If
    '11
    Dim devclock As DATETIMEINFO
    
    If (CKT_GetDeviceClock(IDNumber, devclock)) = 1 Then
        msg$ = "Clock: " & devclock.Year & "-" & devclock.Month & "-" & devclock.Day & Chr$(10) & _
               "       " & devclock.Hour & ":" & devclock.Minute & ":" & devclock.Second
        
         Call CKT_Disconnect
         MsgBox msg$
    Else
        MsgBox ("CKT_GetDeviceClock fail")
    End If
    
End Sub

Private Sub Command65_Click()
If CKT_SetComTimeouts(10 * 1000) Then
    MsgBox "CKT_SetComTimeouts success."
Else
    MsgBox "CKT_SetComTimeouts fail."
End If
End Sub

Private Sub Command66_Click()
Dim msg As CKT_MessageInfo
    msg.PersonID = Int(Text3.Text)
    msg.Year1 = 2011
    msg.Month1 = 2
    msg.Day1 = 1
    msg.Year2 = 2012
    msg.Month2 = 12
    msg.Day2 = 30
            
    msg.msg = "anviz" + Text3.Text
            If (CKT_AddMessage(IDNumber, msg)) = 1 Then
                MsgBox ("CKT_AddMessage OK!")
            End If
End Sub

Private Sub Command67_Click()
Dim pst(28) As Byte
Dim ret As Long
Dim i As Long
For i = 0 To 28
    pst(i) = 0
Next

ret = CKT_GetStateChangeInfo(IDNumber, 1, pst(0))
If ret = 1 Then
    MsgBox ("CKT_GetStateChangeInfo OK")
Else
    MsgBox ("CKT_GetStateChangeInfo fail")
End If
    
End Sub

Private Sub Command68_Click()
Dim pst(28) As Byte
Dim ret As Long
Dim i As Long
For i = 0 To 28
    pst(i) = i + 1
Next

ret = CKT_SetStateChangeInfo(IDNumber, 1, pst(0))
If ret = 1 Then
    MsgBox ("CKT_SetStateChangeInfo OK")
Else
    MsgBox ("CKT_SetStateChangeInfo fail")
End If
End Sub

Private Sub Command69_Click()
Dim ret As Long
Dim Num As Long
Dim N As Long
Dim i As Long
Dim j As Long
Dim Ret2 As Long
Dim str As String
Dim RecordRec(19) As CKT_PictureFileHead
Dim RetCount As Long
Dim InputID As Long
Dim InputFileName As String
Num = 0
N = 0
ret = CKT_GetPictureFileHead(IDNumber, 1, 20, RecordRec(0), RetCount)
If ret = 1 Then
    If RetCount < 1 Then Exit Sub
    
    Do While RetCount > 0
        Num = Num + RetCount
        For i = 0 To RetCount - 1
            N = N + 1
            InputID = RecordRec(i).id
            str = ""
            For j = 0 To 19
                str = str & Chr(RecordRec(i).stime(j))
            Next
            InputFileName = str
        Next
        ret = CKT_GetPictureFileHead(IDNumber, 0, 20, RecordRec(0), RetCount)
        If ret <> 1 Then
            MsgBox ("CKT_GetPictureFileHead fail")
            Exit Sub
        End If
    Loop
    MsgBox ("CKT_GetPictureFileHead OK")
    
    ret = CKT_GetPictureFile(IDNumber, InputID, InputFileName, "1.jpg")
    If ret = 1 Then
        MsgBox ("CKT_GetPictureFile OK")
    Else
        MsgBox ("CKT_GetPictureFile Fail")
    End If
    
    ret = CKT_DelPictureFile(IDNumber, InputID, InputFileName, "1.jpg")
    If ret = 1 Then
        MsgBox ("CKT_DelPictureFile OK")
    Else
        MsgBox ("CKT_DelPictureFile Fail")
    End If
Else
    MsgBox ("CKT_GetPictureFileHead fail")
End If
End Sub

Private Sub Command70_Click()
If Timer2.Enabled Then
   Timer2.Enabled = False
   Else
   Timer2.Enabled = True
   End If
End Sub

Private Sub Command71_Click()
Dim ckqs As GPRSinfo
Dim i As Integer
Dim porti As Long

If (CKT_GetGPRS(IDNumber, ckqs)) = 1 Then
    msg$ = Trim(Left(ckqs.GGSN, InStr(ckqs.GGSN, Chr(0)) - 1)) & Chr$(10)
    msg$ = msg$ + "ServerIP: " & ckqs.ServerIP(0) & "." & ckqs.ServerIP(1) & "." & ckqs.ServerIP(2) & "." & ckqs.ServerIP(3) & Chr$(10)
    porti = HEX_to_DEC(DEC_to_HEX(ckqs.Port(0)) & DEC_to_HEX(ckqs.Port(1)))
    msg$ = msg$ & porti
    MsgBox msg$

    MsgBox ("CKT_GetGPRS OK")
Else
    MsgBox ("CKT_GetGPRS fail")
End If
ckqs.Port(0) = ckqs.Port(1) + 1
ckqs.ServerIP(3) = ckqs.ServerIP(3) + 1
If (CKT_SetGPRS(IDNumber, ckqs)) = 1 Then
        MsgBox ("CKT_SetGPRS OK")
    Else
        MsgBox ("CKT_SetGPRS fail")
    End If

End Sub

Private Sub Form_Load()
    Combo1.ListIndex = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    CKT_Disconnect
    Sleep (500)
End Sub

Private Sub Command1_Click()

If (CKT_ForceOpenLock(IDNumber)) = 1 Then
        MsgBox ("CKT_ForceOpenLock OK")
    Else
        MsgBox ("CKT_ForceOpenLock fail")
    End If
    
End Sub

Private Sub Command2_Click()
    Dim devnetinfo As NETINFO
    
    If CKT_GetDeviceNetInfo(IDNumber, devnetinfo) Then
        msg$ = "IP: " & devnetinfo.IP(0) & "." & devnetinfo.IP(1) & "." & devnetinfo.IP(2) & "." & devnetinfo.IP(3) & Chr$(10)
        msg$ = msg$ + "Mask: " & devnetinfo.Mask(0) & "." & devnetinfo.Mask(1) & "." & devnetinfo.Mask(2) & "." & devnetinfo.Mask(3) & Chr$(10)
        msg$ = msg$ + "Gate: " & devnetinfo.Gateway(0) & "." & devnetinfo.Gateway(1) & "." & devnetinfo.Gateway(2) & "." & devnetinfo.Gateway(3) & Chr$(10)
        msg$ = msg$ + "Server: " & devnetinfo.ServerIP(0) & "." & devnetinfo.ServerIP(1) & "." & devnetinfo.ServerIP(2) & "." & devnetinfo.ServerIP(3) & Chr$(10)
        msg$ = msg$ + "MAC: " & devnetinfo.MAC(0) & "." & devnetinfo.MAC(1) & "." & devnetinfo.MAC(2) & "." & devnetinfo.MAC(3) & "." & devnetinfo.MAC(4) & "." & devnetinfo.MAC(5) & Chr$(10)
        MsgBox msg$
    End If
End Sub

Private Sub Command3_Click()
    Dim IP(3) As Byte
    
    IP(0) = 192
    IP(1) = 168
    IP(2) = 0
    IP(3) = 148
    
    If CKT_SetDeviceIPAddr(IDNumber, IP(0)) Then
        msg$ = "New IP: 192.168.10.254"
    Else
        msg$ = "Fail to set IP to (192.168.10.254)"
    End If
    
    MsgBox msg$
End Sub

Private Sub Command4_Click()
    Dim Mask(3) As Byte
    
    Mask(0) = 255
    Mask(1) = 255
    Mask(2) = 255
    Mask(3) = 0
    
    If CKT_SetDeviceMask(IDNumber, Mask(0)) Then
        msg$ = "New Mask: 255.255.255.0"
    Else
        msg$ = "Fail to set Mask to (255.255.255.0)"
    End If
    
    MsgBox msg$
End Sub

Private Sub Command5_Click()
    Dim Gate(3) As Byte
    
    Gate(0) = 192
    Gate(1) = 168
    Gate(2) = 10
    Gate(3) = 1
    
    If CKT_SetDeviceGateway(IDNumber, Gate(0)) Then
        msg$ = "New Gate: 192.168.10.1"
    Else
        msg$ = "Fail to set Gate to (192.168.10.1)"
    End If
    
    MsgBox msg$
End Sub

Private Sub Command6_Click()
    Dim SvrIP(3) As Byte
    
    SvrIP(0) = 192
    SvrIP(1) = 168
    SvrIP(2) = 10
    SvrIP(3) = 2
    
    If CKT_SetDeviceServerIPAddr(IDNumber, SvrIP(0)) Then
        msg$ = "New SvrIP: 192.168.10.2"
    Else
        msg$ = "Fail to set SvrIP to (192.168.10.2)"
    End If
    
    MsgBox msg$
End Sub

Private Sub Command7_Click()
    Dim MAC(5) As Byte
    
    MAC(0) = 160
    MAC(1) = 168
    MAC(2) = 10
    MAC(3) = 2
    MAC(4) = 10
    MAC(5) = 2
    
    If CKT_SetDeviceMAC(IDNumber, MAC(0)) Then
        msg$ = "New MAC: 160-168-10-2-10-2"
    Else
        msg$ = "Fail to set MAC to (160-168-10-2-10-2)"
    End If
    
    MsgBox msg$
End Sub

Private Sub Command8_Click()
    Dim devclock As DATETIMEINFO
    
    If (CKT_GetDeviceClock(IDNumber, devclock)) = 1 Then
        msg$ = "Clock: " & devclock.Year & "-" & devclock.Month & "-" & devclock.Day & Chr$(10) & _
               "       " & devclock.Hour & ":" & devclock.Minute & ":" & devclock.Second
        MsgBox msg$
    Else
        MsgBox ("CKT_GetDeviceClock fail")
    End If
End Sub

Private Sub Command9_Click()
    Dim tim As SYSTEMTIME
    GetLocalTime tim
    
    If CKT_SetDeviceDate(IDNumber, tim.wYear, tim.wMonth, tim.wDay) Then
        MsgBox "Sucess to send date"
    End If
    
    Sleep (300)
    
    GetLocalTime tim
    
    If CKT_SetDeviceTime(IDNumber, tim.wHour, tim.wMinute, tim.wSecond) Then
        MsgBox "Sucess to send time"
    End If
End Sub

Private Sub Command10_Click()
    Dim pFPData As Long
    Dim FPDataLen As Long
    Dim vbFPData() As Byte
    
    If CKT_GetFPTemplate(IDNumber, 2, 0, pFPData, FPDataLen) = 1 Then
        ReDim vbFPData(FPDataLen - 1) As Byte
        Call PCopyMemory(vbFPData(0), pFPData, FPDataLen)
        CKT_FreeMemory (pFPData)
        
        ' now there are fingerprint data in vbFPData .
        i = 0
        For Each By In vbFPData
            If i = 10 Then
                msg$ = msg$ + Chr(10)
                i = 0
            End If
            msg$ = msg$ + Hex(By) & " "
            i = i + 1
        Next 'i
        
        MsgBox msg$
        
        If CKT_PutFPTemplate(IDNumber, 2, 1, vbFPData(0), FPDataLen) = 1 Then
            MsgBox " the first fingerprint wrote and the second fingerprint successful"
        Else
            MsgBox "the first fingerprint wrote and the second fingerprint failed "
        End If
    End If
End Sub

Private Sub Command11_Click()
    If CKT_GetFPTemplateSaveFile(IDNumber, 2, 0, "C:\1.anv") = 1 Then
        MsgBox "fingerprint data save in C:\1.anv"
    End If
End Sub

Private Sub Command12_Click()
    If CKT_PutFPTemplateLoadFile(IDNumber, 2, 0, "C:\1.anv") = 1 Then
        MsgBox "fingerprint data get from C:\1.anv£¬and download to time&attendance device"
    End If
End Sub

Private Sub Command13_Click()
    Dim i As Integer
    Dim ret As Integer
    Dim PsNo As Long
    Dim tempP As Long
    ret = CKT_ReportConnections(PsNo)
    If ret > 0 Then
        For i = 0 To ret - 1
            Call PCopyMemory(tempP, PsNo, 4) 'long(Integer(PsNo)+I*Sizeof(Integer));
            PsNo = PsNo + 4
            msg$ = msg$ & Hex(tempP) & " "
        Next
        MsgBox msg$
        If CKT_FreeMemory(PsNo) Then
            MsgBox "CKT_ReportConnections ok."
        End If
    End If
End Sub

Private Sub Command14_Click()
    Dim mpiRet As Long
    Dim person As PERSONINFO
    
    With person
        .CardNo = &H1000000
        .Name = "zyp"
        .Password(0) = &H31
        .Password(1) = &H32
        .Password(2) = &H33
        .Password(3) = &H34
        .Password(4) = 0
        .PersonID = 1
        .Dept = 0
        .FPMark = 0
        .Group = 0
        .KQOption = 6
        .Other = 0
    End With
    
    mpiRet = CKT_ModifyPersonInfo(IDNumber, person)
    If mpiRet = CKT_RESULT_ADDOK Then
        MsgBox "add successfully"
    ElseIf mpiRet = CKT_RESULT_CHANGEOK Then
        MsgBox "modeify successfully "
    ElseIf mpiRet = CKT_ERROR_MEMORYFULL Then
        MsgBox "memory full"
    Else
        MsgBox "communication failed"
    End If
End Sub

Private Sub Command15_Click()
    Dim dpiRet As Long
    dpiRet = CKT_DeletePersonInfo(IDNumber, 5, 255)
    If dpiRet = CKT_RESULT_OK Then
        MsgBox "delete successfully"
    ElseIf dpiRet = CKT_ERROR_NOTHISPERSON Then
        MsgBox "user ID not exist"
    Else
        MsgBox "communication failed"
    End If
End Sub

Private Sub Command16_Click()
    If CKT_DeleteAllPersonInfo(IDNumber) Then
        MsgBox "delete all users data OK"
    Else
        MsgBox "delete all users data fail"
    End If
End Sub

Private Sub Command17_Click()
    Dim Section(337) As Byte
    Dim ret As Long
    Dim ret1 As Long
    
    ret = CKT_GetFPRawData(IDNumber, 2, 0, Section(0))
    If ret = CKT_RESULT_OK Then
        For Each By In Section
            msg$ = msg$ & Hex(By) & " "
        Next
        MsgBox msg$
        ret1 = CKT_PutFPRawData(IDNumber, 2, 0, Section(0), 338)
        If ret1 = 1 Then
            MsgBox " the first fingerprint wrote and the second fingerprint successful"
        Else
            MsgBox "the first fingerprint wrote and the second fingerprint failed "
        End If
        
    ElseIf ret = CKT_ERROR_NOTHISPERSON Then
        MsgBox "user ID not exist"
    End If
End Sub

Private Sub Command19_Click()
    Dim RecordCount, RetCount As Long
    Dim pPersons, pLongRun As Long
    Dim person As PERSONINFO
    Dim X As Integer
    
    If CKT_ListPersonInfoEx(IDNumber, pLongRun) Then
        Do While True
            ret = CKT_ListPersonProgress(pLongRun, RecordCount, RetCount, pPersons)
            If ret = 0 Then
               Exit Sub
            End If
            
            If (ret <> 0) Then
                Dim ptemp As Long
                ptemp = pPersons
                
                For i = 0 To RetCount - 1
                    Call PCopyMemory(person, pPersons, PERSONINFOSIZE)
                    X = InStr(person.Name, Chr(0)) - 1
                    If X > 0 Then
                    person.Name = Left(person.Name, X)
                    End If
                    pPersons = pPersons + PERSONINFOSIZE
                    
                    msg$ = msg$ & person.PersonID & "  " & Trim(person.Name) & ","
                    If i Mod 10 = 9 Then
                        MsgBox "RetCount:" & RetCount & Chr(10) & msg$
                        msg$ = ""
                    End If
                Next
                
                If msg$ <> "" Then
                    MsgBox msg$
                End If
                
                If ptemp <> 0 Then
                    Call CKT_FreeMemory(ptemp)
                End If
            End If
            
            If ret <> 2 Then
                Exit Sub
            End If
        Loop
    End If

End Sub

Private Sub Command20_Click()
    Dim personCount, FPCount, clockingCount As Long
    If CKT_GetCounts(IDNumber, personCount, FPCount, clockingCount) Then 'CKT_GetCounts
        msg$ = "Person: " & personCount & Chr(10) & "Finger Prints: " & FPCount & Chr(10) & "Clocking Record: " & clockingCount
        MsgBox msg$
    Else
        MsgBox "CKT_GetCounts failed"
    End If
End Sub



Private Sub Command22_Click()
    If CKT_ClearClockingRecord(IDNumber, 0, 0) Then
        MsgBox "clear off all records"
    Else
        MsgBox "communication failed"
    End If
End Sub

Private Sub Command23_Click()
If (CKT_SetRepeatKQ(IDNumber, 1)) = 1 Then
        MsgBox ("CKT_SetRepeatKQ OK")
    Else
        MsgBox ("CKT_SetRepeatKQ fail")
    End If
End Sub

Private Sub Command24_Click()
    Dim RecordCount, RetCount As Long
    Dim pClockings, pLongRun As Long
    Dim clocking As CLOCKINGRECORD
    
    If CKT_GetClockingNewRecordEx(IDNumber, pLongRun) Then
        Do While True
            ret = CKT_GetClockingRecordProgress(pLongRun, RecordCount, RetCount, pClockings)
            If ret = 0 Then
               Exit Sub
            End If
            
            If (ret <> 0) Then
                Dim ptemp As Long
                ptemp = pClockings
                
                For i = 1 To RetCount
                    Call PCopyMemory(clocking, pClockings, CLOCKINGRECORDSIZE)
                    pClockings = pClockings + CLOCKINGRECORDSIZE
                    
                    msg$ = msg$ & clocking.PersonID & "  "
                    'If i Mod 50 = 0 Then
                    '    MsgBox "RetCount:" & RetCount & Chr(10) & msg$
                    '    msg$ = ""
                    'End If
                Next
                
                If msg$ <> "" Then
                    MsgBox msg$
                End If
                
                If ptemp <> 0 Then
                    Call CKT_FreeMemory(ptemp)
                End If
            End If
            
            If ret = 1 Then
                Exit Sub
            End If
        Loop
    End If
End Sub

Private Sub Command25_Click()
If (CKT_SetDeviceMode(IDNumber, 0)) = 1 Then
        MsgBox ("CKT_SetDeviceMode OK")
    Else
        MsgBox ("CKT_SetDeviceMode fail")
    End If
End Sub

Private Sub Command26_Click()
    If CKT_SetDoor(IDNumber, 2) Then
        MsgBox "door opening 2 seconds"
    End If
End Sub

Private Sub Command27_Click()
    If CKT_SetSpeakerVolume(IDNumber, 20) Then
        MsgBox "the maximum volume of speaker"
    End If
End Sub

Private Sub Command28_Click()
    If CKT_SetDeviceAdminPassword(IDNumber, "9999") Then
        MsgBox "modify admin password"
    End If
End Sub

Private Sub Command29_Click()
        If CKT_SetRealtimeMode(IDNumber, 1) Then
            
            MsgBox "enable realtime supervision mode"
        End If
End Sub

Private Sub Timer2_Timer()
'    End
'Exit Sub

    Dim count As Long
    Dim pClockings, ptemp As Long
    Dim clocking As CLOCKINGRECORD
    
    count = CKT_ReadRealtimeClocking(pClockings)
    
    ptemp = pClockings
    For i = 1 To count
        Call PCopyMemory(clocking, ptemp, CLOCKINGRECORDSIZE)
        ptemp = ptemp + CLOCKINGRECORDSIZE
        
        msg$ = msg$ & clocking.PersonID & "  "
    Next
    
    If msg$ <> "" Then
        MsgBox msg$
    End If
    
    If pClockings <> 0 Then
        CKT_FreeMemory (pClockings)
    End If
End Sub

Private Sub Command30_Click()
    If CKT_ResetDevice(IDNumber) Then
        MsgBox "reset system successfully"
    End If
End Sub

Private Sub Command31_Click()
    Dim devnfo As DEVICEINFO
        
    If CKT_GetDeviceInfo(IDNumber, devnfo) Then
        MsgBox devnfo.id & Chr(10) & _
            devnfo.MajorVersion & "." & devnfo.MinorVersion & Chr(10) & _
            devnfo.SpeakerVolume
    End If
End Sub

Private Sub Command32_Click()
    If CKT_SetWG(IDNumber, 1) Then
        MsgBox "set Wiegand as ANVIZ32 successfully"
    End If
End Sub


