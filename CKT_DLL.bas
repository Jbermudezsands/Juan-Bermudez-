Attribute VB_Name = "CKT_DLL"
' Consts

Public Const CKT_ERROR_INVPARAM = -1
Public Const CKT_ERROR_NETDAEMONREADY = -1
Public Const CKT_ERROR_NOTHISPERSON = -1
Public Const CKT_ERROR_CHECKSUMERR = -2
Public Const CKT_ERROR_MEMORYFULL = -1
Public Const CKT_ERROR_INVFILENAME = -3
Public Const CKT_ERROR_FILECANNOTOPEN = -4
Public Const CKT_ERROR_FILECONTENTBAD = -5
Public Const CKT_ERROR_FILECANNOTCREATED = -2
Public Const CKT_RESULT_OK = 1
Public Const CKT_RESULT_ADDOK = 1
Public Const CKT_RESULT_CHANGEOK = 2
Public Const CKT_RESULT_HASMORECONTENT = 2


Public Const PERSONINFOSIZE = 48
Public Const PERSONINFOSIZEEX = 100
Public Const CLOCKINGRECORDSIZE = 40

' Types
Public Type CKT_KQState
    Num As Long
    kqmsg(9, 15) As Byte
End Type

Public Type CKT_PictureFileHead
    id As Long
    stime(19) As Byte
End Type

Public Type MessageHead '机器时间
    PersonID As Long               '考勤机序列号
    sYear As Long
    sMon As Long
    sDay As Long
    eYear As Long
    eMon As Long
    eDay As Long
End Type

Public Type RingTime '打铃时间
    Hour   As Long
    Minute As Long
    Week   As Long
End Type

Public Type TimeSectEX '时段时间
    z1(3)   As Byte
    z2(3)   As Byte
    z3(3)   As Byte
    z4(3)   As Byte
    z5(3)   As Byte
    z6(3)   As Byte
    z7(3)   As Byte
End Type


Public Type TimeSect '时段时间
    bHour   As Long
    bMinute As Long
    eHour   As Long
    eMinute As Long
End Type

Public Type NETINFO
   id          As Long
   IP(3)       As Byte
   Mask(3)     As Byte
   Gateway(3)  As Byte
   ServerIP(3) As Byte
   MAC(5)      As Byte
End Type

Public Type DATETIMEINFO
   id     As Long
   Year   As Integer
   Month  As Byte
   Day    As Byte
   Hour   As Byte
   Minute As Byte
   Second As Byte
End Type

Public Type PERSONINFO
   PersonID As Long
   Password(7) As Byte
   CardNo   As Long
   Name     As String * 12
   Dept     As Long
   Group    As Long
   KQOption As Long
   FPMark   As Long
   Other    As Long
End Type

Public Type PERSONINFOEX
   PersonID As Long
   Password(7) As Byte
   CardNo   As Long
   Name     As String * 64
   Dept     As Long
   Group    As Long
   KQOption As Long
   FPMark   As Long
   Other    As Long
End Type

Public Type CLOCKINGRECORD
   id       As Long
   PersonID As Long
   Stat     As Long
   BackupCode     As Long
   WorkTyte     As Long
   Time     As String * 20
End Type


Public Type DEVICEINFO
    id              As Long
    MajorVersion    As Long
    MinorVersion    As Long
    SpeakerVolume   As Long
    Parameter       As Long
    DefaultAuth     As Long
    FixWGHead       As Long
    WGOption        As Long
    AutoUpdateAllow As Long
    KQRepeatTime    As Long
    RealTimeAllow   As Long
    RingAllow       As Long
    DoorLockDelay   As Long
    AdminPassword   As String * 8
End Type


Public Type CKT_MessageInfo
   PersonID As Long
   Year1   As Long
   Month1  As Long
   Day1    As Long
   Year2   As Long
   Month2  As Long
   Day2    As Long
   msg     As String * 48

End Type

Public Type GPRSinfo
    GGSN As String * 16
    ServerIP(3) As Byte
    Port(1) As Byte
End Type

Public Type CKT_Daylight
   X0 As Byte
   X1 As Byte
   X2 As Byte
   X3 As Byte
   X4 As Byte
   X5 As Byte
   X6 As Byte
   X7 As Byte
   X8 As Byte
   X9 As Byte
   X10 As Byte
   X11 As Byte
   X12 As Byte
   X13 As Byte
   X14 As Byte
   X15 As Byte
End Type

' Routines
Public Declare Function CKT_GetPictureFileHead Lib "tc400.dll" (ByVal Sno As Long, ByVal param As Long, ByVal reqcount As Long, ByRef Xls As CKT_PictureFileHead, ByRef RetCount As Long) As Long
Public Declare Function CKT_GetPictureFile Lib "tc400.dll" (ByVal Sno As Long, ByVal usrID As Long, ByVal phead As String, ByVal pfile As String) As Long
Public Declare Function CKT_DelPictureFile Lib "tc400.dll" (ByVal Sno As Long, ByVal usrID As Long, ByVal phead As String, ByVal pfile As String) As Long


Public Declare Function CKT_GetStateChangeInfo Lib "tc400.dll" (ByVal Sno As Long, ByVal ord As Long, ByRef Xls As Byte) As Long
Public Declare Function CKT_SetStateChangeInfo Lib "tc400.dll" (ByVal Sno As Long, ByVal ord As Long, ByRef Xls As Byte) As Long

Public Declare Function CKT_GetDaylightSavingTime Lib "tc400.dll" (ByVal Sno As Long, ByRef Xls As Byte) As Long
Public Declare Function CKT_SetDaylightSavingTime Lib "tc400.dll" (ByVal Sno As Long, ByRef Xls As Byte) As Long

Public Declare Function CKT_SetNetTimeouts Lib "tc400.dll" (ByVal LTime As Long) As Long
Public Declare Function CKT_SetComTimeouts Lib "tc400.dll" (ByVal LTime As Long) As Long

Public Declare Sub CKT_EnableLog Lib "tc400.dll" (f As Long)

Public Declare Function CKT_GetMachineNumber Lib "tc400.dll" (ByVal Sno As Long, ByVal Num As String) As Long

Public Declare Function CKT_FreeMemory Lib "tc400.dll" (ByVal memory As Long) As Long

Public Declare Function CKT_ForceOpenLock Lib "tc400.dll" (ByVal Sno As Long) As Long

Public Declare Function CKT_SetDeviceMode Lib "tc400.dll" (ByVal Sno As Long, ByVal Mode As Long) As Long
Public Declare Function CKT_SetRingAllow Lib "tc400.dll" (ByVal Sno As Long, ByVal tpe As Long) As Long
Public Declare Function CKT_SetAutoUpdate Lib "tc400.dll" (ByVal Sno As Long, ByVal au As Long) As Long
Public Declare Function CKT_SetRepeatKQ Lib "tc400.dll" (ByVal Sno As Long, ByVal timea As Long) As Long

Public Declare Function CKT_GetTimeSection Lib "tc400.dll" (ByVal Sno As Long, ByVal ord As Long, ByRef pts As TimeSectEX) As Long
Public Declare Function CKT_SetTimeSection Lib "tc400.dll" (ByVal Sno As Long, ByVal ord As Long, pts As TimeSect) As Long

Public Declare Function CKT_GetGroup Lib "tc400.dll" (ByVal Sno As Long, ByVal ord As Long, ByRef pts As Long) As Long
Public Declare Function CKT_SetGroup Lib "tc400.dll" (ByVal Sno As Long, ByVal ord As Long, ByRef pts As Long) As Long

Public Declare Function CKT_ChangeConnectionMode Lib "tc400.dll" (ByVal Mode As Long) As Long

Public Declare Function CKT_SetWorkCode Lib "tc400.dll" (ByVal Sno As Long, ByVal Mode As Long) As Long

Public Declare Function CKT_GetHitRingInfo Lib "tc400.dll" (ByVal Sno As Long, ByRef prt As RingTime) As Long
Public Declare Function CKT_SetHitRingInfo Lib "tc400.dll" (ByVal Sno As Long, ByVal ord As Long, ByRef prt As RingTime) As Long

Public Declare Function CKT_GetKQState Lib "tc400.dll" (ByVal Sno As Long, ByRef kqs As CKT_KQState) As Long
Public Declare Function CKT_SetKQState Lib "tc400.dll" (ByVal Sno As Long, ByRef kqs As CKT_KQState) As Long

Public Declare Function CKT_SetDateTimeFormat Lib "tc400.dll" (ByVal Sno As Long, ByVal dateF As Long, ByVal timeF As Long) As Long

Public Declare Function CKT_RegisterSno Lib "tc400.dll" (ByVal Sno As Long, ByVal ComPort As Long) As Long
Public Declare Function CKT_RegisterNet Lib "tc400.dll" (ByVal Sno As Long, ByVal Addr As String) As Long
Public Declare Function CKT_RegisterUSB Lib "tc400.dll" (ByVal Sno As Long, ByVal index As Long) As Long

Public Declare Sub CKT_UnregisterSnoNet Lib "tc400.dll" (ByVal Sno As Long)
Public Declare Function CKT_NetDaemon Lib "tc400.dll" () As Long
Public Declare Function CKT_NetDaemonWithPort Lib "tc400.dll" (ByVal Portint As Long) As Long
Public Declare Function CKT_ComDaemon Lib "tc400.dll" () As Long
Public Declare Sub CKT_Disconnect Lib "tc400.dll" ()

Public Declare Function CKT_ReportConnections Lib "tc400.dll" (ByRef ppSno As Long) As Long


Public Declare Function CKT_ModifyDeviceSno Lib "tc400.dll" (ByVal Sno As Long, ByVal Num As Long) As Long
Public Declare Function CKT_SetSleepTime Lib "tc400.dll" (ByVal Sno As Long, ByVal min As Long) As Long


Public Declare Function CKT_GetDeviceNetInfo Lib "tc400.dll" (ByVal Sno As Long, pNetInfo As NETINFO) As Long
Public Declare Function CKT_SetDeviceIPAddr Lib "tc400.dll" (ByVal Sno As Long, ByRef IP As Byte) As Long
Public Declare Function CKT_SetDeviceMask Lib "tc400.dll" (ByVal Sno As Long, ByRef Mask As Byte) As Long
Public Declare Function CKT_SetDeviceGateway Lib "tc400.dll" (ByVal Sno As Long, ByRef Gate As Byte) As Long
Public Declare Function CKT_SetDeviceServerIPAddr Lib "tc400.dll" (ByVal Sno As Long, ByRef Svr As Byte) As Long
Public Declare Function CKT_SetDeviceMAC Lib "tc400.dll" (ByVal Sno As Long, ByRef MAC As Byte) As Long


Public Declare Function CKT_GetDeviceClock Lib "tc400.dll" (ByVal Sno As Long, pDateTimeInfo As DATETIMEINFO) As Long
Public Declare Function CKT_SetDeviceClock Lib "tc400.dll" (ByVal Sno As Long, pDateTimeInfo As DATETIMEINFO) As Long
Public Declare Function CKT_SetDeviceDate Lib "tc400.dll" (ByVal Sno As Long, ByVal Year As Integer, ByVal Month As Byte, ByVal Day As Byte) As Long
Public Declare Function CKT_SetDeviceTime Lib "tc400.dll" (ByVal Sno As Long, ByVal Hour As Byte, ByVal Minute As Byte, ByVal Second As Byte) As Long


Public Declare Function CKT_GetFPTemplate Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByRef pFPData As Long, ByRef FPDataLen As Long) As Long
Public Declare Function CKT_PutFPTemplate Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByRef pFPData As Byte, ByVal FPDataLen As Long) As Long
Public Declare Function CKT_GetFPTemplateSaveFile Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByVal FPDataFilename As String) As Long
Public Declare Function CKT_PutFPTemplateLoadFile Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByVal FPDataFilename As String) As Long


Public Declare Function CKT_GetFPRawData Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByRef FPRawData As Byte) As Long
Public Declare Function CKT_PutFPRawData Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByRef FPRawData As Byte, ByVal FPDataLen As Long) As Long
Public Declare Function CKT_GetFPRawDataSaveFile Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByVal FPDataFilename As String) As Long
Public Declare Function CKT_PutFPRawDataLoadFile Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal FPID As Long, ByVal FPDataFilename As String) As Long


Public Declare Function CKT_ListPersonInfo Lib "tc400.dll" (ByVal Sno As Long, ByRef pRecordCount As Long, ByRef ppPersons As Long) As Long
Public Declare Function CKT_ModifyPersonInfo Lib "tc400.dll" (ByVal Sno As Long, person As PERSONINFO) As Long
Public Declare Function CKT_ModifyPersonInfoLongName Lib "tc400.dll" (ByVal Sno As Long, person As PERSONINFOEX) As Long
Public Declare Function CKT_DeletePersonInfo Lib "tc400.dll" (ByVal Sno As Long, ByVal PersonID As Long, ByVal backupID As Long) As Long
Public Declare Function CKT_DeleteAllPersonInfo Lib "tc400.dll" (ByVal Sno As Long) As Long


Public Declare Function CKT_ListPersonInfoEx Lib "tc400.dll" (ByVal Sno As Long, ByRef ppLongRun As Long) As Long
Public Declare Function CKT_ListPersonProgress Lib "tc400.dll" (ByVal pLongRun As Long, ByRef pRecCount As Long, ByRef pRetCount As Long, ByRef ppPersons As Long) As Long
Public Declare Function CKT_ListPersonProgressLongName Lib "tc400.dll" (ByVal pLongRun As Long, ByRef pRecCount As Long, ByRef pRetCount As Long, ByRef ppPersons As Long) As Long


Public Declare Function CKT_GetCounts Lib "tc400.dll" (ByVal Sno As Long, ByRef pPersonCount As Long, ByRef pFPCount As Long, ByRef pClockingsCount As Long) As Long
Public Declare Function CKT_GetClockingRecord Lib "tc400.dll" (ByVal Sno As Long, ByRef pRecordCount As Long, ByRef ppClockings As Long) As Long
Public Declare Function CKT_ClearClockingRecord Lib "tc400.dll" (ByVal Sno As Long, ByVal tpe As Long, ByVal count As Long) As Long
Public Declare Function CKT_RecallClockingRecord Lib "tc400.dll" (ByVal Sno As Long, ByVal NewRecordCount As Long) As Long


Public Declare Function CKT_GetClockingRecordEx Lib "tc400.dll" (ByVal Sno As Long, ByRef ppLongRun As Long) As Long
Public Declare Function CKT_GetClockingNewRecordEx Lib "tc400.dll" (ByVal Sno As Long, ByRef ppLongRun As Long) As Long
Public Declare Function CKT_GetClockingRecordProgress Lib "tc400.dll" (ByVal pLongRun As Long, ByRef pRecCount As Long, ByRef pRetCount As Long, ByRef ppPersons As Long) As Long


Public Declare Function CKT_SetDoor Lib "tc400.dll" (ByVal Sno As Long, ByVal Second As Long) As Long
Public Declare Function CKT_SetSpeakerVolume Lib "tc400.dll" (ByVal Sno As Long, ByVal Volume As Long) As Long
Public Declare Function CKT_SetDeviceAdminPassword Lib "tc400.dll" (ByVal Sno As Long, ByVal Password As String) As Long

Public Declare Function CKT_SetRealtimeMode Lib "tc400.dll" (ByVal Sno As Long, ByVal RealMode As Long) As Long
Public Declare Function CKT_SetWG Lib "tc400.dll" (ByVal Sno As Long, ByVal WGMode As Long) As Long
Public Declare Function CKT_SetNoSearch Lib "tc400.dll" (ByVal Sno As Long, ByVal dis_n_search As Long) As Long

                        
Public Declare Function CKT_ReadRealtimeClocking Lib "tc400.dll" (ByRef ppClockings As Long) As Long


Public Declare Function CKT_ResetDevice Lib "tc400.dll" (ByVal Sno As Long) As Long
Public Declare Function CKT_GetDeviceInfo Lib "tc400.dll" (ByVal Sno As Long, devinfo As DEVICEINFO) As Long

Public Declare Function CKT_GetMessageByIndex Lib "tc400.dll" (ByVal Sno As Long, ByVal idx As Integer, ByRef msg As CKT_MessageInfo) As Integer
Public Declare Function CKT_AddMessage Lib "tc400.dll" (ByVal Sno As Long, ByRef msg As CKT_MessageInfo) As Integer
Public Declare Function CKT_DelMessageByIndex Lib "tc400.dll" (ByVal Sno As Long, ByVal idx As Integer) As Integer
Public Declare Function CKT_GetAllMessageHead Lib "tc400.dll" (ByVal Sno As Long, ByRef amh As MessageHead) As Long
'GPRSinfo
Public Declare Function CKT_GetGPRS Lib "tc400.dll" (ByVal Sno As Long, ByRef kqs As GPRSinfo) As Long
Public Declare Function CKT_SetGPRS Lib "tc400.dll" (ByVal Sno As Long, ByRef kqs As GPRSinfo) As Long
