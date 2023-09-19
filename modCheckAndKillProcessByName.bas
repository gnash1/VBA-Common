Attribute VB_Name = "modCheckAndKillProcessByName"
Option Explicit
'File:   modCheckAndKillProcessByName
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        4/5/10

'KillProcessByNameAll - Process Kill for All OS's - must have psapi.dll
'Two methods are included below.
'A simplified method called "KillProcessByNameWMI" implemented at the end of this file 9/12/2007

Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32.dll" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Function TerminateProcess Lib "kernel32.dll" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long

Private Const PROCESS_TERMINATE = &H1 'enables terminate process in NT
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
'Private Const MAX_PATH = 260

'KillProcessByNameNew -
'These functions are not supported on NT 3&4 and require Win 9x or NT 5 (2000) or higher.
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long
'Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Private Const TH32CS_SNAPPROCESS As Long = 2& 'Not included in NT 4.0
Private Const MAX_PATH As Long = 260

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type

'For checking availabilty of Procs
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

'Two methods used here, one old, one new.
Public Function KillProcessByName(ByVal strApp As String) As Boolean
    Dim retval As Boolean
    'Check to see what methods are available before proceeding
'    If APIFunctionPresent("CreateToolhelp32Snapshot", "kernel32.dll") And _
'      APIFunctionPresent("Process32First", "kernel32.dll") And _
'      APIFunctionPresent("Process32Next", "kernel32.dll") Then
        retval = KillProcessByNameWMI(strApp, True)
'        retval = KillProcessByNameNew(strApp, True)
'    Else
'        retval = KillProcessByNameAll(strApp)
'    End If
    KillProcessByName = retval
End Function

Public Function IsProcessRunning(ByVal strApp As String) As Boolean
    IsProcessRunning = KillProcessByNameWMI(strApp, False)
End Function

'Kills the first occurance of an application in the process list
'Requires psapi.dll to be installed on the host computer.
'The code below is a NT (3& 4) version of Kill Process code:
'http://www.experts-exchange.com/Programming/Programming_Languages/Visual_Basic/Q_10259330.html
'(if you don't have psapi.dll in your NT, copy this file from your Visual Studio CD : \UNSUPPRT\WSVIEW\WINNT\ to c:\winnt\system32\)
Private Function KillProcessByNameAll(ByVal strApp As String) As Boolean
On Error GoTo errsub:
    Dim hProcess As Long
    Dim cb As Long
    Dim cbNeeded As Long
    Dim NumElements As Long
    Dim ProcessIDs() As Long
    Dim cbNeeded2 As Long
    Dim NumElements2 As Long
    Dim Modules(1 To 200) As Long
    Dim lRet As Long 'Return Values
    Dim ModuleName As String
    Dim nSize As Long
    Dim i As Long
    Dim ExCode As Long
    'Dim strApp As String
   
    'Get the array containing the process id's for each process object
    cb = 8
    cbNeeded = 96
    Do While cb <= cbNeeded
        cb = cb * 2
        ReDim ProcessIDs(cb / 4) As Long
        lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
    Loop
    NumElements = cbNeeded / 4

    For i = 1 To NumElements
        'Get a handle to the Process
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, ProcessIDs(i))
        'Got a Process handle
        If hProcess <> 0 Then
            'Get an array of the module handles for the specified process
            lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
            'If the Module Array is retrieved, Get the ModuleFileName
            If lRet <> 0 Then
                ModuleName = Space(MAX_PATH)
                nSize = 500
                lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
                ModuleName = Trim(ModuleName)
                strApp = Trim(strApp)
                ModuleName = PathParse(ModuleName, 2)
                ModuleName = Left(ModuleName, Len(strApp))
                If UCase(ModuleName) = UCase(strApp) Then
                'If InStr(1, UCase(Trim(ModuleName)), UCase(Trim(strApp))) <> 0 Then
                    
                    lRet = GetExitCodeProcess(hProcess, ExCode)
                    lRet = TerminateProcess(hProcess, ExCode) 'Kill the process
                    'Close the handle to the process
                    lRet = CloseHandle(hProcess)
                    KillProcessByNameAll = True
                    'Exit Function
                End If
            End If
        End If
   'Close the handle to the process
    lRet = CloseHandle(hProcess)
    Next
    KillProcessByNameAll = False
    Exit Function
errsub:
KillProcessByNameAll = False
End Function

'Kills all occurances of an application in the process list if bKillProcess is True
'If bKillProcess is false, returns if process is running, without killing it.
'Supported on Win9x and WinNT 5.0 (2000) and higher.
Private Function KillProcessByNameNew(ByVal strAppName As String, bKillProcess As Boolean) As Boolean
On Error GoTo errsub:
    Dim hProcess As Long 'handle to process
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim lProcess As Long
    Dim strProcName As String
    Dim lRet As Long 'Return Values
    Dim ExCode As Long
    Dim colKillList As New Collection
    Dim i As Integer
    
    hSnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = -1 Then
        GoTo errsub
    End If
    
    uProcess.dwSize = Len(uProcess)
    lProcess = ProcessFirst(hSnapShot, uProcess)
    
    'Loop through processes
    Do While lProcess
        strProcName = Trim(uProcess.szExeFile)
        
        If InStr(1, LCase(strProcName), LCase(strAppName)) Then
            colKillList.Add uProcess.th32ProcessID
            KillProcessByNameNew = True
        End If
        lProcess = ProcessNext(hSnapShot, uProcess)
    Loop
    
    'Close the handles
    lRet = CloseHandle(hSnapShot)
    
    'kill all the found process's
    If bKillProcess Then
        For i = 1 To colKillList.Count
            hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, colKillList(i))
            lRet = GetExitCodeProcess(hProcess, ExCode)
            lRet = TerminateProcess(hProcess, ExCode) 'Kill the process
            
            'Close the handles
            lRet = CloseHandle(hProcess)
            lRet = CloseHandle(hSnapShot)
        Next i
    End If
    Exit Function
    
errsub:
KillProcessByNameNew = False
End Function

'Parses Path and returns according to Mode setting: 1=Path / 2=Filename
Private Function PathParse(mPath As String, Mode As Integer)
    Dim CurPos, TempPos
    CurPos = 1
    Do
        TempPos = InStr(CurPos, mPath, "\")
        If TempPos <> 0 Then CurPos = TempPos + 1
    Loop Until TempPos = 0
    If Mode = 1 Then 'Return Path
        PathParse = Mid(mPath, 1, CurPos - 1)
    ElseIf Mode = 2 Then 'Return Filename
        PathParse = Mid(mPath, CurPos)
    End If
End Function

Private Function APIFunctionPresent(ByVal FunctionName As String, ByVal DllName As String) As Boolean
   'http://www.freevbcode.com/ShowCode.Asp?ID=429
   On Error GoTo errsub
    Dim lHandle As Long
    Dim lAddr  As Long
    lHandle = GetModuleHandle(DllName)
    lHandle = LoadLibrary(DllName) 'Doesn't work for kernel32.dll
    If lHandle <> 0 Then
        lAddr = GetProcAddress(lHandle, FunctionName)
        FreeLibrary lHandle
    End If
    APIFunctionPresent = (lAddr <> 0)
errsub:
APIFunctionPresent = False
End Function

Private Function KillProcessByNameWMI(procName As String, bKillProcess As Boolean) As Boolean
'Kills all occurances of an application in the process list if bKillProcess is True
'If bKillProcess is false, returns if process is running, without killing it.
    On Error GoTo 0
    Dim objProcList As Object 'SWbernObjectSet
    Dim objWMI As Object 'SWbernServicesEx
    Dim objProc As Object 'SWbernObjectEx
    
    'create WMI object instance
    Set objWMI = GetObject("winmgmts:")
        If Not IsNull(objWMI) Then
        'create object collection of Win32 processes
        Set objProcList = objWMI.InstancesOf("win32_process")
        For Each objProc In objProcList 'iterate through enumerated collection
            If UCase(objProc.name) = UCase(procName) Then
                KillProcessByNameWMI = True
                If bKillProcess = True Then
                    Debug.Print "Terminating Process " & procName
                    objProc.Terminate (0)
                    Debug.Print procName & " was terminated"
                End If
            End If
        Next
    End If
    Set objProcList = Nothing
    Set objWMI = Nothing
End Function

Public Function KillProcessByHwndApplicationSimple(hWndApplication As Long) As Boolean
'Used to kill a particular instance of an application by PID (processor ID).
'Modified from:
'http://www.microsoft.com/technet/scriptcenter/resources/qanda/sept04/hey0927.mspx
'Uses: Windows Management Instrumentation - Microsoft WMI Scripting v1.2 Scripting
    On Error GoTo 0
    Dim objProcList As Object 'SWbernObjectSet
    Dim objWMI As Object 'SWbernServicesEx
    Dim objProc As Object 'SWbernObjectEx
    Dim lRet As Long
    Dim lProcessID As Long
    
    lRet = GetWindowThreadProcessId(hWndApplication, lProcessID)
        
    'create WMI object instance
    Set objWMI = GetObject("winmgmts:")
    
    If Not IsNull(objWMI) And lRet > 0 Then
        Set objProcList = objWMI.ExecQuery("Select * from Win32_Process Where ProcessID = " & lProcessID & "")
        
        Debug.Print "Cleaning up Excel.exe with Process ID: " & lProcessID
        
        For Each objProc In objProcList
            objProc.Terminate (0)
        Next
    End If
    
    Set objProc = Nothing
    Set objProcList = Nothing
    Set objWMI = Nothing
End Function
