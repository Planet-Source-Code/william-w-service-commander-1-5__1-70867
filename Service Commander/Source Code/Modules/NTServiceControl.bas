Attribute VB_Name = "NTServiceControl"
'http://www.microsoft.com/msj/0298/service.aspx
Option Explicit
'2008 Bilgus
'------------TYPE DECLARES----------------

Public Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion(1 To 128) As Byte      '  Maintenance string for PSS usage
End Type

Private Type QUERY_SERVICE_CONFIG
   dwServiceType As Long
   dwStartType As Long
   dwErrorControl As Long
   lpBinaryPathName As Long
   lpLoadOrderGroup As Long
   dwTagId As Long
   lpDependencies As Long
   lpServiceStartName As Long
   lpDisplayName As Long
End Type

Private Type SERVICE_STATUS
   dwServiceType As Long
   dwCurrentState As Long
   dwControlsAccepted As Long
   dwWin32ExitCode As Long
   dwServiceSpecificExitCode As Long
   dwCheckPoint As Long
   dwWaitHint As Long
   dwProcessId As Long
   dwServiceFlags As Long
End Type

Public Type SvcReturn
   Error As Long
   Account As String
   DisplayName As String
   Dependencies As String
   ErrorControl As String
   TagId As String
   LoadOrderGroup As String
   PathName As String
   StartType As String
   ServiceType As String
End Type

'---------------------CONSTANTS--------------------------
Private Const SERVICE_ACCEPT_STOP = &H1, SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Private Const SERVICE_ACCEPT_SHUTDOWN = &H4, SERVICE_STATE_ALL As Long = &H3

Private Const SC_ENUM_PROCESS_INFO As Long = 0
Private Const GENERIC_READ As Long = &H80000000

Private Const SC_MANAGER_CONNECT = &H1&, SC_MANAGER_CREATE_SERVICE = &H2&
Private Const SC_MANAGER_ENUMERATE_SERVICE = &H4, SC_MANAGER_LOCK = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS = &H10, SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Private Const SC_MANAGER_ALL_ACCESS = SC_MANAGER_CONNECT + SC_MANAGER_CREATE_SERVICE + _
   SC_MANAGER_ENUMERATE_SERVICE + SC_MANAGER_LOCK + SC_MANAGER_QUERY_LOCK_STATUS + _
   SC_MANAGER_MODIFY_BOOT_CONFIG

Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_QUERY_CONFIG = &H1&, SERVICE_CHANGE_CONFIG = &H2&
Private Const SERVICE_QUERY_STATUS = &H4&, SERVICE_ENUMERATE_DEPENDENTS = &H8&
Private Const SERVICE_START = &H10&, SERVICE_STOP = &H20&
Private Const SERVICE_PAUSE_CONTINUE = &H40&, SERVICE_INTERROGATE = &H80&
Private Const SERVICE_USER_DEFINED_CONTROL = &H100&
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED + SERVICE_QUERY_CONFIG + _
   SERVICE_CHANGE_CONFIG + SERVICE_QUERY_STATUS + SERVICE_ENUMERATE_DEPENDENTS + SERVICE_START + _
   SERVICE_STOP + SERVICE_PAUSE_CONTINUE + SERVICE_INTERROGATE + SERVICE_USER_DEFINED_CONTROL)

Private Const SERVICE_ACTIVE = &H1, SERVICE_INACTIVE = &H2

Private Const ERROR_PATH_NOT_FOUND = 3
Private Const ERROR_ACCESS_DENIED = 5, ERROR_INVALID_HANDLE = 6
Private Const ERROR_INVALID_PARAMETER = 87, ERROR_INSUFFICIENT_BUFFER = 122
Private Const ERROR_INVALID_NAME = 123, ERROR_MORE_DATA = 234

Private Const ERROR_DEPENDENT_SERVICES_RUNNING = 1051, ERROR_INVALID_SERVICE_CONTROL = 1052
Private Const ERROR_SERVICE_REQUEST_TIMEOUT = 1053, ERROR_SERVICE_NO_THREAD = 1054
Private Const ERROR_DATABASE_LOCKED = 1055, ERROR_SERVICE_ALREADY_RUNNING = 1056
Private Const ERROR_INVALID_SERVICE_ACCOUNT = 1057, ERROR_SERVICE_DISABLED = 1058
Private Const ERROR_CIRCULAR_DEPENDENCY = 1059, ERROR_SERVICE_DOES_NOT_EXIST = 1060
Private Const ERROR_SERVICE_CANNOT_ACCEPT_CONTROL = 1061, ERROR_SERVICE_NOT_ACTIVE = 1062
Private Const ERROR_FAILED_SERVICE_CONTROLLER_CONNECT = 1063, ERROR_EXCEPTION_IN_SERVICE = 1064
Private Const ERROR_DATABASE_DOES_NOT_EXIST = 1065, ERROR_SERVICE_SPECIFIC_ERROR = 1066
Private Const ERROR_PROCESS_ABORTED = 1067, ERROR_SERVICE_DEPENDENCY_FAIL = 1068
Private Const ERROR_SERVICE_LOGON_FAILED = 1069, ERROR_SERVICE_START_HANG = 1070
Private Const ERROR_INVALID_SERVICE_LOCK = 1071, ERROR_SERVICE_MARKED_FOR_DELETE = 1072
Private Const ERROR_SERVICE_EXISTS = 1073, ERROR_ALREADY_RUNNING_LKG = 1074
Private Const ERROR_SERVICE_DEPENDENCY_DELETED = 1075, ERROR_BOOT_ALREADY_ACCEPTED = 1076
Private Const ERROR_SERVICE_NEVER_STARTED = 1077, ERROR_DUPLICATE_SERVICE_NAME = 1078
Private Const ERROR_DIFFERENT_SERVICE_ACCOUNT = 1079, ERROR_SERVICE_NOT_FOUND = 1243
Private Const ERROR_SERVICE_CODE_OWNER = "066073076071085083"

Public Const SERVICE_NO_CHANGE = &HFFFFFFFF
Public Const SERVICE_ERROR_IGNORE = &H0, SERVICE_ERROR_NORMAL = &H1
Public Const SERVICE_ERROR_SEVERE = &H2, SERVICE_ERROR_CRITICAL = &H3

Public Const SERVICE_BOOT_START = &H0, SERVICE_SYSTEM_START = &H1
Public Const SERVICE_AUTO_START = &H2, SERVICE_DEMAND_START = &H3, SERVICE_DISABLED = &H4

Public Const SERVICE_KERNEL_DRIVER = &H1, SERVICE_FILE_SYSTEM_DRIVER = &H2
Public Const SERVICE_WIN32_OWN_PROCESS = &H10
Public Const SERVICE_WIN32_SHARE_PROCESS = &H20, SERVICE_INTERACTIVE_PROCESS = &H100
Public Const SERVICE_WIN32 = &H30, SERVICE_DRIVER As Long = &HB
Public Const VER_PLATFORM_WIN32_NT = 2&
Public Const Rev = 1.5
'-------------------ENUM DECLARES------------------------

Private Type ENUM_SERVICE_STATUS
   lpServiceName As Long
   lpDisplayName As Long
   ServiceStatus As SERVICE_STATUS
End Type

Private Enum SERVICE_CONTROL
   SERVICE_CONTROL_STOP = 1
   SERVICE_CONTROL_PAUSE = 2
   SERVICE_CONTROL_CONTINUE = 3
   SERVICE_CONTROL_INTERROGATE = 4
   SERVICE_CONTROL_SHUTDOWN = 5
End Enum

Public Enum SERVICE_STATE
   SERVICE_STOPPED = &H1
   SERVICE_START_PENDING = &H2
   SERVICE_STOP_PENDING = &H3
   SERVICE_RUNNING = &H4
   SERVICE_CONTINUE_PENDING = &H5
   SERVICE_PAUSE_PENDING = &H6
   SERVICE_PAUSED = &H7
End Enum

'-----------------API DECLARES---------------------
'C
Public Declare Function ChangeServiceConfig Lib "advapi32.dll" _
      Alias "ChangeServiceConfigA" ( _
      ByVal hService As Long, _
      ByVal dwServiceType As Long, _
      ByVal dwStartType As Long, _
      ByVal dwErrorControl As Long, _
      ByVal lpBinaryPathName As Long, _
      ByVal lpLoadOrderGroup As Long, _
      ByVal lpdwTagId As Long, _
      ByVal lpDependencies As Long, _
      ByVal lpServiceStartName As Long, _
      ByVal lpPassword As Long, _
      ByVal lpDisplayName As Long) As Long

Private Declare Function CloseServiceHandle Lib "advapi32" (ByVal hSCObject As Long) As Long

Private Declare Function ControlService Lib "advapi32" ( _
      ByVal hService As Long, _
      ByVal dwControl As SERVICE_CONTROL, _
      lpServiceStatus As SERVICE_STATUS) As Long

Private Declare Sub CopyMemory Lib "kernel32" _
      Alias "RtlMoveMemory" ( _
      Destination As Any, _
      Source As Any, _
      ByVal Length As Long)

'D
Private Declare Function DeleteService Lib "advapi32" (ByVal hService As Long) As Long

'E
Private Declare Function EnumServicesStatusEx Lib "advapi32.dll" _
      Alias "EnumServicesStatusExA" ( _
      ByVal hSCManager As Long, _
      ByVal InfoLevel As Long, _
      ByVal dwServiceType As Long, _
      ByVal dwServiceState As Long, _
      lpServices As Long, _
      ByVal cbBufSize As Long, _
      pcbBytesNeeded As Long, _
      ByRef lpServicesReturned As Long, _
      lpResumeHandle As Long, _
      pszGroupName As Long) As Long

'F
Private Declare Function FormatMessage Lib "kernel32" _
      Alias "FormatMessageA" ( _
      ByVal dwFlags As Long, _
      lpSource As Any, _
      ByVal dwMessageId As Long, _
      ByVal dwLanguageId As Long, _
      ByVal lpBuffer As String, _
      ByVal nSize As Long, _
      Arguments As Long) As Long

'G
Private Declare Function GetVersionEx Lib "kernel32" _
      Alias "GetVersionExA" ( _
      lpVersionInformation As OSVERSIONINFO) As Long

'L
Private Declare Function lstrcpy Lib "kernel32" _
      Alias "lstrcpyA" ( _
      ByVal lpString1 As Any, _
      ByVal lpString2 As Any) As Long

Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long

'O
Private Declare Function OpenSCManager Lib "advapi32" _
      Alias "OpenSCManagerA" ( _
      ByVal lpMachineName As String, _
      ByVal lpDatabaseName As String, _
      ByVal dwDesiredAccess As Long) As Long

Private Declare Function OpenService Lib "advapi32" _
      Alias "OpenServiceA" ( _
      ByVal hSCManager As Long, _
      ByVal lpServiceName As String, _
      ByVal dwDesiredAccess As Long) As Long

'Q
Private Declare Function QueryServiceConfig Lib "advapi32" _
      Alias "QueryServiceConfigA" ( _
      ByVal hService As Long, _
      lpServiceConfig As Any, _
      ByVal cbBufSize As Long, _
      pcbBytesNeeded As Long) As Long

Private Declare Function QueryServiceStatus Lib "advapi32" ( _
      ByVal hService As Long, _
      lpServiceStatus As SERVICE_STATUS) As Long

'S
Private Declare Function SetServiceStatus Lib "advapi32" ( _
      ByVal hService As Long, _
      lpServiceStatus As SERVICE_STATUS) As Long

Private Declare Function StartService Lib "advapi32" _
      Alias "StartServiceA" ( _
      ByVal hService As Long, _
      ByVal dwNumServiceArgs As Long, _
      ByVal lpServiceArgVectors As Long) As Long
'---------------------Subs and Functions--------------------


Public Function CheckIsNT() As Boolean

   ' CheckIsNT() returns True, if the OS is NT based
   ' and False otherwise.

  Dim OSVer As OSVERSIONINFO

   OSVer.dwOSVersionInfoSize = LenB(OSVer)
   GetVersionEx OSVer
   CheckIsNT = OSVer.dwPlatformId = VER_PLATFORM_WIN32_NT

End Function

Public Function DeleteNTService(ServiceName As String) As Long

   ' This function uninstalls service
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long
  Dim Status As SERVICE_STATUS

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_ALL_ACCESS)
      'open selected service to get or set desired state

      If hService <> 0 Then
         ' Stop service if it is running
         ControlService hService, SERVICE_CONTROL_STOP, Status

         If DeleteService(hService) = 0 Then
            DeleteNTService = Err.LastDllError
         End If

         CloseServiceHandle hService
       Else
         DeleteNTService = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      DeleteNTService = Err.LastDllError
   End If

End Function

Public Sub EnumServices(ListBxType As ListBox, _
                        ListBxName As ListBox, _
                        ListBxDesc As ListBox, _
                        ListBxPath As ListBox, _
                        ListBxState As ListBox, _
                        ListBxStart As ListBox, _
                        Optional Delimeter As Long = 0, _
                        Optional AltDelimeter As Long = 0, _
                        Optional ShowService As Long = SERVICE_DRIVER Or SERVICE_WIN32, _
                        Optional ShowState As Long = SERVICE_STATE_ALL)

   'Fills ListBoxes Defined above with service names and service configurations

  Dim a As Long
  Dim API As String
  Dim bMore As Boolean
  Dim bFail As Boolean
  Dim CurSvc As SvcReturn
  Dim Data() As Byte
  Dim DisplayName As String
  Dim EnumResult As Long
  Dim hSCManager As Long
  Dim lBytesNeeded As Long
  Dim lBytesExtra As Long
  Dim lNumberOfServices As Long
  Dim lStart As Long
  Dim ServiceName As String
  Dim sService As ENUM_SERVICE_STATUS
  

 
   hSCManager = OpenSCManager(vbNullString, vbNullString, GENERIC_READ)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      'if we didn't get an error then

      Do
         bMore = False
         ' Obtain the size of the buffer required so we make a call to it first
         EnumResult = EnumServicesStatusEx(hSCManager, SC_ENUM_PROCESS_INFO, ShowService, ShowState, _
            0, 0, lBytesNeeded, lNumberOfServices, lStart, 0)

         If Err.LastDllError = ERROR_INSUFFICIENT_BUFFER Or Err.LastDllError = ERROR_MORE_DATA Then
            ReDim Data(lBytesNeeded - 1)
            'see if our buffer was to small or if we didn't get all the data
            EnumResult = EnumServicesStatusEx(hSCManager, SC_ENUM_PROCESS_INFO, ShowService, _
               ShowState, ByVal VarPtr(Data(0)), lBytesNeeded, lBytesExtra, lNumberOfServices, _
               lStart, ByVal CLng(0))
            ' Check if there are more services to fetch and set a flag if there are

            If EnumResult = 0 And Err.LastDllError = ERROR_MORE_DATA Then
               bMore = True
               EnumResult = 1
               Err.Clear
            End If

            If EnumResult <> 0 Then
               If lNumberOfServices <> 0 Then
                  a = 0

                  Do
                     CopyMemory sService, Data(a * Len(sService)), Len(sService)
                     'copy the pointers into the sservice structure
                     'we can't process the pointers directly
                     ServiceName = PtrToString(sService.lpServiceName)
                     'Get the data the pointer is pointing to
                     CurSvc = GetServiceConfig(ServiceName)
                     ' If there is no Service Name check for system and Idle processes

                     If ServiceName = "" Then
                        If sService.ServiceStatus.dwProcessId = 0 Then ServiceName = "System Idle"
                        If sService.ServiceStatus.dwProcessId = 8 Then ServiceName = "System"
                     End If

                     DisplayName = CurSvc.DisplayName

                     If Len(DisplayName) = 0 Then DisplayName = "NONE"

                     If ServiceName <> "UNKNOWN" Then
                        'place the data into our lists defined at the call of the function

                        If CurSvc.ServiceType = Delimeter Or CurSvc.ServiceType = AltDelimeter Or _
                           Delimeter = 0 Then
                           ListBxType.AddItem SvcType(CurSvc.ServiceType)
                           ListBxName.AddItem ServiceName
                           ListBxDesc.AddItem CurSvc.DisplayName
                           ListBxPath.AddItem CurSvc.PathName
                           ListBxState.AddItem SvcState(GetServiceStatus(ServiceName))
                           ListBxStart.AddItem StartType(CurSvc.StartType)
                          
                        End If

                        'SERVICE_STOPPED = &H1
                        'SERVICE_START_PENDING = &H2
                        'SERVICE_STOP_PENDING = &H3
                        'SERVICE_RUNNING = &H4
                        'SERVICE_CONTINUE_PENDING = &H5
                        'SERVICE_PAUSE_PENDING = &H6
                        'SERVICE_PAUSED = &H7
                     End If

                     a = a + 1
                  Loop Until a = lNumberOfServices Or bFail = True

                  'continue until we have no services left or there is an error
                Else
                  bMore = False
               End If

             Else
               bFail = True
               API = "EnumServices Failure"
            End If

          Else
            bFail = False
            API = "EnumServiceStatus - Get Buffer Size Failed"
         End If

      Loop Until bMore = False Or bFail = True
    Else
      bFail = True
      API = "OpenSCManager Failed"
   End If

   If bFail = True Then
      MsgBox "API Call Failed " & API
   End If

   CloseServiceHandle hSCManager

End Sub

Public Function ErrLib(Error As Long) As String

   'makes Error number into Error text

   Select Case Error
    Case 0: ErrLib = "NO_ERROR"
    Case ERROR_DATABASE_LOCKED: ErrLib = "ERROR_DATABASE_LOCKED"
    Case Else:
      Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100, FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
      Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_STRING = &H400
      Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
      Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
      Dim sBuff As String
      Dim lCount As Long

      ' Return the error message associated with Error:
      sBuff = String$(256, 0)
      lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, Error, _
         0&, sBuff, Len(sBuff), ByVal 0)

      If lCount Then
         ErrLib = Left$(sBuff, lCount)
       Else
         ErrLib = "UNKNOWN_ERROR " & Error
      End If

   End Select

End Function

Public Function GetServiceConfig(ServiceName As String) As SvcReturn

   'returns Svcreturn Structure with service configuration

  Dim hSCManager As Long
  Dim hService As Long
  Dim SCfg() As QUERY_SERVICE_CONFIG
  Dim lBuffer As Long
  Dim lBytesNeeded As Long
  Dim lStructNeeded As Long

   'Dim s As String
   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_QUERY_CONFIG)
      'open selected service to get or set desired state
      Call QueryServiceConfig(hService, ByVal &H0, &H0, lBytesNeeded)

      If Err.LastDllError <> ERROR_INSUFFICIENT_BUFFER Then
         GetServiceConfig.Error = Err.LastDllError
         Exit Function
      End If

      'Calculate the buffer sizes
      lStructNeeded = lBytesNeeded / Len(SCfg(0)) + 1

      ReDim SCfg(lStructNeeded - 1)
      lBuffer = lStructNeeded * Len(SCfg(0))

      Call QueryServiceConfig(hService, SCfg(0), lBuffer, lBytesNeeded)

      With GetServiceConfig
         .Account = PtrToString(SCfg(0).lpServiceStartName)
         .Dependencies = PtrToString(SCfg(0).lpDependencies)
         .DisplayName = PtrToString(SCfg(0).lpDisplayName)
         .ErrorControl = SCfg(0).dwErrorControl
         .LoadOrderGroup = PtrToString(SCfg(0).lpLoadOrderGroup)
         .PathName = PtrToString(SCfg(0).lpBinaryPathName)
         .ServiceType = SCfg(0).dwServiceType
         .StartType = SCfg(0).dwStartType
         .TagId = SCfg(0).dwTagId
      End With

      CloseServiceHandle hService
      CloseServiceHandle hSCManager
    Else
      GetServiceConfig.Error = Err.LastDllError
   End If

End Function

Public Function GetServiceStatus(ServiceName As String) As SERVICE_STATE

   ' This function Retrieves the service status
   ' (Running,stopped, started, starting pausing, paused ,resuming, running...)
  Dim hSCManager As Long
  Dim hService As Long
  Dim Status As SERVICE_STATUS

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_QUERY_STATUS)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If QueryServiceStatus(hService, Status) Then
            GetServiceStatus = Status.dwCurrentState
         End If

         CloseServiceHandle hService
      End If

      CloseServiceHandle hSCManager
   End If

End Function

Public Function PauseNTService(ServiceName As String) As Long

   ' This function Pauses service
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long
  Dim Status As SERVICE_STATUS

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_PAUSE_CONTINUE)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If ControlService(hService, SERVICE_CONTROL_PAUSE, Status) = 0 Then
            PauseNTService = Err.LastDllError
         End If

         CloseServiceHandle hService
       Else
         PauseNTService = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      PauseNTService = Err.LastDllError
   End If

End Function

Public Function PtrToString(Pointer As Long) As String

   'turns a pointer into the string it points to
   PtrToString = Space(255)
   lstrcpy PtrToString, Pointer
   PtrToString = Left$(PtrToString, lstrlen(PtrToString))

End Function

Public Function ResumeNTService(ServiceName As String) As Long

   ' This function Resumes service
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long
  Dim Status As SERVICE_STATUS

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_PAUSE_CONTINUE)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If ControlService(hService, SERVICE_CONTROL_CONTINUE, Status) = 0 Then
            ResumeNTService = Err.LastDllError
         End If

         CloseServiceHandle hService
       Else
         ResumeNTService = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      ResumeNTService = Err.LastDllError
   End If

End Function

Public Function SetServiceConfig(ServiceName As String, _
                                 ServiceType As Long, _
                                 StartType As Long, _
                                 ErrorControl As Long, _
                                 TagId As Long) As Long

   'http://msdn.microsoft.com/en-us/library/ms681987(VS.85).aspx
   'This Function Sets a new service configuration
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_CHANGE_CONFIG)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If ChangeServiceConfig(hService, ServiceType, StartType, ErrorControl, 0&, 0&, TagId, 0&, _
            0&, 0&, 0&) = 0 Then SetServiceConfig = Err.LastDllError
         CloseServiceHandle hService
       Else
         SetServiceConfig = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      SetServiceConfig = Err.LastDllError
   End If

End Function

Public Function StartNTService(ServiceName As String) As Long

   ' This function starts service
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_START)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If StartService(hService, 0, 0) = 0 Then
            StartNTService = Err.LastDllError
         End If

         CloseServiceHandle hService
       Else
         StartNTService = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      StartNTService = Err.LastDllError
   End If

End Function

Public Function StartType(Stype As Variant) As String

   ' Get A String start type from a number

   Select Case Stype
    Case SERVICE_BOOT_START
      StartType = "Boot"

    Case SERVICE_SYSTEM_START
      StartType = "System"

    Case SERVICE_AUTO_START
      StartType = "Automatic"

    Case SERVICE_DISABLED
      StartType = "Disabled"

    Case SERVICE_DEMAND_START
      StartType = "Manual"

    Case Else
      StartType = "Unknown"
   End Select

End Function

Public Function StopNTService(ServiceName As String) As Long

   ' This function stops service
   ' Returns nonzero value on error

  Dim hSCManager As Long
  Dim hService As Long
  Dim Status As SERVICE_STATUS

   hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
   'Open the service manager with desired access

   If hSCManager <> 0 Then
      hService = OpenService(hSCManager, ServiceName, SERVICE_STOP)
      'open selected service to get or set desired state

      If hService <> 0 Then
         If ControlService(hService, SERVICE_CONTROL_STOP, Status) = 0 Then
            StopNTService = Err.LastDllError
         End If

         CloseServiceHandle hService
       Else
         StopNTService = Err.LastDllError
      End If

      CloseServiceHandle hSCManager
    Else
      StopNTService = Err.LastDllError
   End If

End Function

Public Function SvcError(SState As Variant) As String

   'Gets Errorstate string from number
   'The severity of the error, and action taken
   'if the service fails to start

   Select Case SState
    Case 0:
      SvcError = "Ignore Error"

    Case 1:
      SvcError = "Normal Error"

    Case 2:
      SvcError = "Sever Error"

    Case 3:
      SvcError = "Critical Error"

    Case Else:
      SvcError = "Unknown? " & SState
      'Debug.Print SvcError & " Error Control"
   End Select

End Function

Public Function SvcState(SState As Variant) As String

   'Gets service state string from number

   Select Case SState
    Case 1:
      SvcState = "Stopped"

    Case 2:
      SvcState = "Starting"

    Case 3:
      SvcState = "Stopping"

    Case 4:
      SvcState = "Running"

    Case 5:
      SvcState = "Continuing"

    Case 6:
      SvcState = "Pausing"

    Case 7:
      SvcState = "Paused"

    Case Else:
      SvcState = "Unknown? " & SState
      'Debug.Print SvcState & " Service State"
   End Select

End Function

Public Function SvcType(Stype As Variant) As String

   'Gets Service Type String from number

   Select Case Stype
    Case 1:
      SvcType = "Kernel Mode Driver"

    Case 2:
      SvcType = "File System Driver"

    Case 4:
      SvcType = "Adapter"

    Case 8:
      SvcType = "Driver Service" 'File System Driver Service or Reconizer Driver

    Case 10, 16:
      SvcType = "Win32 Own Process"

    Case 20, 32:
      SvcType = "Win32 Shared Process"

    Case 100, 272:
      SvcType = "Own Interactive Process"

    Case 120, 288:
      SvcType = "Shared Interactive"

    Case Else:
      SvcType = "Unknown? " & Stype & " Service Type"
      'Debug.Print SvcType & " Service Type"
   End Select

End Function

