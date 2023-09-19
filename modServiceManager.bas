Attribute VB_Name = "modServiceManager"
'File:   modServiceManager
'Author:      Greg Harward
'Date:        12/7/08
'
'Summary:
'A module which will start/stop and query NT services, passing the ServerName as a valid remote machine name will control the service on that machine, "" as the servername will do local services.
'Top portion is Win32 code.  Bottom portion is a simpler form using WMI scripts.
'
'Online References:
'http://www.experts-exchange.com/Programming/Programming_Languages/Visual_Basic/Q_20187550.html
'Revisions:
'Date     Initials    Description of changes

Option Explicit

Private Declare Function OpenService Lib "ADVAPI32.DLL" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenSCManager Lib "ADVAPI32.DLL" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function QueryServiceStatus Lib "ADVAPI32.DLL" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function ControlService Lib "ADVAPI32.DLL" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function StartService Lib "ADVAPI32.DLL" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private Declare Function CloseServiceHandle Lib "ADVAPI32.DLL" (ByVal hSCObject As Long) As Long

Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const STANDARD_RIGHTS_ALL = &H1F0000
Const SPECIFIC_RIGHTS_ALL = &HFFFF

' Service database names
Const SERVICES_ACTIVE_DATABASE = "ServicesActive"
Const SERVICES_FAILED_DATABASE = "ServicesFailed"

' Value to indicate no change to an optional parameter
Const SERVICE_NO_CHANGE = &HFFFF

' Service State -- for Enum Requests (Bit Mask)
Const SERVICE_ACTIVE = &H1
Const SERVICE_INACTIVE = &H2
Const SERVICE_STATE_ALL = (SERVICE_ACTIVE Or SERVICE_INACTIVE)

' Controls
Const SERVICE_CONTROL_STOP = &H1
Const SERVICE_CONTROL_PAUSE = &H2
Const SERVICE_CONTROL_CONTINUE = &H3
Const SERVICE_CONTROL_INTERROGATE = &H4
Const SERVICE_CONTROL_SHUTDOWN = &H5

' Service State -- for CurrentState
Const SERVICE_STOPPED = &H1
Const SERVICE_START_PENDING = &H2
Const SERVICE_STOP_PENDING = &H3
Const SERVICE_RUNNING = &H4
Const SERVICE_CONTINUE_PENDING = &H5
Const SERVICE_PAUSE_PENDING = &H6
Const SERVICE_PAUSED = &H7

' Controls Accepted  (Bit Mask)
Const SERVICE_ACCEPT_STOP = &H1
Const SERVICE_ACCEPT_PAUSE_CONTINUE = &H2
Const SERVICE_ACCEPT_SHUTDOWN = &H4

' Service Control Manager object specific access types
Const SC_MANAGER_CONNECT = &H1
Const SC_MANAGER_CREATE_SERVICE = &H2
Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Const SC_MANAGER_LOCK = &H8
Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20

Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)

' Service object specific access type
Const SERVICE_QUERY_CONFIG = &H1
Const SERVICE_CHANGE_CONFIG = &H2
Const SERVICE_QUERY_STATUS = &H4
Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Const SERVICE_START = &H10
Const SERVICE_STOP = &H20
Const SERVICE_PAUSE_CONTINUE = &H40
Const SERVICE_INTERROGATE = &H80
Const SERVICE_USER_DEFINED_CONTROL = &H100
Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Public Type SERVICE_STATUS
   dwServiceType As Long
   dwCurrentState As Long
   dwControlsAccepted As Long
   dwWin32ExitCode As Long
   dwServiceSpecificExitCode As Long
   dwCheckPoint As Long
   dwWaitHint As Long
End Type
'strServerName is optional is not passed local machine name is assumed.
Public Function QueryService(strServerName As String, strServiceName As String) As String
Attribute QueryService.VB_ProcData.VB_Invoke_Func = " \n14"
   LogMsg "QueryService : " & strServiceName
   Dim lngSMHandle As Long
   Dim lngSvcHandle As Long
   Dim ssStatus As SERVICE_STATUS
   Dim lRet As Long
   
   lngSMHandle = OpenSCManager(strServerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT)
   lngSvcHandle = OpenService(lngSMHandle, strServiceName, SERVICE_ALL_ACCESS)
   lRet = QueryServiceStatus(lngSvcHandle, ssStatus)
   If lngSvcHandle = 0& Then QueryService = "Service Not Found"
   Select Case ssStatus.dwCurrentState
   Case SERVICE_STOPPED
       QueryService = "Service Is Stopped"
   Case SERVICE_START_PENDING
       QueryService = "Service Start Pending"
   Case SERVICE_STOP_PENDING
       QueryService = "Service Stop Pending"
   Case SERVICE_RUNNING
       QueryService = "Service Is Running"
   Case SERVICE_CONTINUE_PENDING
       QueryService = "Service Continue Pending"
   Case SERVICE_PAUSE_PENDING
       QueryService = "Service Pause Pending"
   Case SERVICE_PAUSED
       QueryService = "Service Paused"
   End Select
   CloseServiceHandle lngSvcHandle
   CloseServiceHandle lngSMHandle
End Function

Public Sub StopSvc(strServerName As String, strServiceName As String)
Attribute StopSvc.VB_ProcData.VB_Invoke_Func = " \n14"
   LogMsg "Stop Service : " & strServiceName
   Dim lngSMHandle As Long
   Dim lngSvcHandle As Long
   Dim ssStatus As SERVICE_STATUS
   Dim lRet As Long
   
   lngSMHandle = OpenSCManager(strServerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT)
   lngSvcHandle = OpenService(lngSMHandle, strServiceName, SERVICE_ALL_ACCESS)
   lRet = QueryServiceStatus(lngSvcHandle, ssStatus)
   If lngSvcHandle <> 0& Then
       ControlService lngSvcHandle, SERVICE_CONTROL_STOP, ssStatus
   End If
   CloseServiceHandle lngSvcHandle
   CloseServiceHandle lngSMHandle
End Sub

Public Sub StartSvc(strServerName As String, strServiceName As String)
Attribute StartSvc.VB_ProcData.VB_Invoke_Func = " \n14"
   LogMsg "Start Service : " & strServiceName
   Dim lngSMHandle As Long
   Dim lngSvcHandle As Long
   Dim ssStatus As SERVICE_STATUS
   Dim lRet As Long
   
   lngSMHandle = OpenSCManager(strServerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_CONNECT)
   lngSvcHandle = OpenService(lngSMHandle, strServiceName, SERVICE_ALL_ACCESS)
   lRet = QueryServiceStatus(lngSvcHandle, ssStatus)
   If lngSvcHandle <> 0& Then
       StartService lngSvcHandle, 0&, 0&
   End If
   CloseServiceHandle lngSvcHandle
   CloseServiceHandle lngSMHandle
End Sub

Public Function StartSvcForTime(strServerName As String, strServiceName As String, intTimoutSecs As Integer) As String
Attribute StartSvcForTime.VB_ProcData.VB_Invoke_Func = " \n14"
    Dim datCurrentTime As Date
    datCurrentTime = Now
    
    'attempt to start service if not running
    Do While QueryService(strServerName, strServiceName) <> "Service Is Running"
        Call StartSvc(strServerName, strServiceName)
        DoEvents
        If DateDiff("s", datCurrentTime, Now) > intTimoutSecs Or QueryService(strServerName, strServiceName) = "Service Is Running" Then
            Exit Do
        End If
    Loop
    StartSvcForTime = QueryService(strServiceName, strServerName)
End Function

Public Sub LogMsg(strMessage)
Attribute LogMsg.VB_ProcData.VB_Invoke_Func = " \n14"
   'Not supported in VBA
   On Error Resume Next
'   app.LogEvent Format(Now(), "DD/MM/YYYY Hh:Mm:Ss : " & strMessage)
   Err.Clear
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WMI Code - Windows Management Instrumentation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function StartServiceForTimeWMI(ByVal strServerName As String, ByVal strServiceName, ByVal lTimeoutSeconds As Long) As Boolean
    On Error GoTo errsub
    Dim dateCurrentTime As Date
    Dim lRet As Long
    
    dateCurrentTime = Now

    'Attempt to start service if not running
    If GetServiceStateByNameWMI(strServerName, strServiceName) <> "Running" Then
        If StartServiceByNameWMI(strServerName, strServiceName) = 0 Then    'Start was successful.
            Do Until GetServiceStateByNameWMI(strServerName, strServiceName) = "Running"
                DoEvents
                If DateDiff("s", dateCurrentTime, Now) > lTimeoutSeconds Then 'Or GetServiceStateByNameWMI(strServerName, strServiceName) = "Running" Then
                    StartServiceForTimeWMI = False
                    Exit Function
                End If
            Loop
        End If
    End If
    StartServiceForTimeWMI = True
    Exit Function
errsub:
    StartServiceForTimeWMI = False
End Function

Private Function StartServiceByNameWMI(ByVal strServerName As String, ByVal strServiceName As String) As Long
'http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/services/#StopService.htm
    Dim objWMIService As Object 'SWbemServicesEx
    Dim objServiceSet As Object 'SWbemObjectSet
'    Dim objService As SWbemObjectSet
    Dim errReturnCode As Long
    Dim obj As Object 'SWbemObjectEx
    
    errReturnCode = -1
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServerName & "\root\cimv2")
    Set objServiceSet = objWMIService.ExecQuery("Select * from win32_Service Where Name = '" & strServiceName & "'") 'One service

    For Each obj In objServiceSet
'        If obj.Name = strServiceName Then
'            Debug.Print obj.Name & " state is: " & obj.State
            If obj.State = "Running" Then
                errReturnCode = 0
            ElseIf obj.State = "Paused" Then
                errReturnCode = obj.ResumeService
            Else
                errReturnCode = obj.StartService    '0 = Success, 10 if already started.    obj.State = "Running"
    '            errReturnCode = obj.StopService     '0 = Success, 5 if already started stopped. obj.State = "Stopped"
    '            errReturnCode = obj.ResumeService   '0 = Success(was paused), 6 = Currently stopped, 10 = was already running.
        'PauseService & ResumeService
            End If
            Exit For
'        End If
    Next obj
    
    StartServiceByNameWMI = errReturnCode
End Function

Private Function GetServiceStateByNameWMI(ByVal strServerName As String, ByVal strServiceName As String) As String
'http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/services/#StopService.htm
    Dim objWMIService As Object 'SWbemServicesEx
    Dim objServiceSet As Object 'SWbemObjectEx
    '    Dim objService As Object 'SWbemObjectSet
    Dim errReturnCode As Long
    Dim obj As Object 'SWbemObjectEx
    
    On Error GoTo errsub
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServerName & "\root\cimv2")
    Set objServiceSet = objWMIService.ExecQuery("Select * from win32_Service Where Name = '" & strServiceName & "'") 'One service
    For Each obj In objServiceSet
'        If obj.Name = strServiceName Then
            GetServiceStateByNameWMI = obj.State
'            Debug.Print obj.Name & " state is: " & obj.State
'            Debug.Print obj.Name & " Status is: " & obj.Status
            Exit Function
'        End If
    Next obj
    Exit Function
errsub:
    GetServiceStateByNameWMI = "Error"
End Function

Private Function GetAllServicesStateByNameWMI(ByVal strServerName As String, ByVal strServiceName As String) As String
'http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/services/#StopService.htm
    Dim objWMIService As Object 'SWbemServicesEx
    Dim objServiceSet As Object 'SWbemObjectSet
'    Dim objService As Object 'SWbemObjectSet
    Dim errReturnCode As Long
    Dim obj As Object 'SWbemObjectEx
    Dim strState As String
    
    strState = "Error"
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServerName & "\root\cimv2")
    Set objServiceSet = objWMIService.ExecQuery("Select * from win32_Service") 'Returns all services
'    Set objService = objWMIService.ExecQuery("Select * from win32_Service Where Name = 'MSSQL$PM2K5'") 'Returns collection containing only one services
    
    For Each obj In objServiceSet
        If obj.Name = strServiceName Then
            strState = obj.State
'            Debug.Print obj.Name & " state is: " & obj.State
            Exit For
        End If
    Next obj
    GetAllServicesStateByNameWMI = strState
End Function
