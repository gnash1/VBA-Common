Attribute VB_Name = "modMainCommon"
Option Explicit

'File:   modMainCommon
'Author:      Greg Harward
'Date:        1/25/2013
'
'Summary:
'Common code used in multiple applications.  Code is organized into sections.
'Contains late binding
'
'Online References:
'
'
'Revisions:
'Date     Initials    Description of changes

'Notes:
'See Excel.Constants for application constants.
'Application.Volatile True - to cause autoupdate when values are modified.

'In Office 2003 these funtions require the 'VBA.' prefix.
'day, hour, minute, dateserial, format, strConv, chr$, chr, Environ, mid, left, right, ucase, string, space, trim, datediff, dateadd, datevalue, Fix, Split, InStr, IsNull

'''''''Directory:
'' File/Directory/Path Functions
'' Control/Object/Collection Functions
'' General Utility Functions
'' Unsorted'

Private Enum eDayOfMonth
    FirstDayOfMonth
    LastDayOfMonth
End Enum

Private Enum nShowCmd
    SW_HIDE = 0 'Hides the window and activates another window.
'    SW_MAXIMIZE = 3 'Maximizes the specified window.
    SW_MINIMIZE = 6 'Minimizes the specified window and activates the next top-level window in the z-order.
    SW_RESTORE = 9 'Activates and displays the window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when restoring a minimized window.
    SW_SHOW = 5 'Activates the window and displays it in its current size and position.
    SW_SHOWDEFAULT = 10 'Sets the show state based on the SW_ flag specified in the STARTUPINFO structure passed to the CreateProcess function by the program that started the application. An application should call ShowWindow with this flag to set the initial show state of its main window.
    SW_SHOWMAXIMIZED = 3 'Activates the window and displays it as a maximized window.
    SW_SHOWMINIMIZED = 2 'Activates the window and displays it as a minimized window.
    SW_SHOWMINNOACTIVE = 7 'Displays the window as a minimized window. The active window remains active.
    SW_SHOWNA = 8 'Displays the window in its current state. The active window remains active.
    SW_SHOWNOACTIVATE = 4 'Displays a window in its most recent size and position. The active window remains active.
    SW_SHOWNORMAL = 1 'Activates and displays a window. If the window is minimized or maximized, Windows restores it to its original size and position. An application should specify this flag when displaying the window for the first time.
End Enum

'Private Enum EnvironmentVariable
'    eALLUSERSPROFILE    'C:\ProgramData
'    eAPPDATA            'C:\Users\(username}\AppData\Roaming
'    eCommonProgramFiles 'C:\Program Files\Common Files
'    eCOMPUTERNAME       '{computername}
'    eCOMSPEC            'C:\Windows\System32\cmd.exe
'    eHOMEDRIVE          'C:
'    eHOMEPATH           '\Users\{username}
'    eLOCALAPPDATA       'C:\Users\{username}\AppData\Local
'    ePATH               'Varies. Includes C:\Windows\System32\;C:\Windows\;C:\Windows\System32\Wbem
'    ePATHEXT            '.COM; .EXE; .BAT; .CMD; .VBS; .VBE; .JS ; .WSF; .WSH; .MSC
'    eProgramData        'C:\ProgramData
'    eProgramFiles       'Directory containing program files, usually C:\Program Files
'    eProgramFiles (x86) 'In 64-bit systems, directory containing 32-bit programs. Usually C:\Program Files (x86)
'    ePROMPT             'Code for current command prompt format. Code is usually $P$G
'    ePublic             'C:\Users\Public
'    eSYSTEMDRIVE        'The drive containing the Windows XP root directory, usually C:
'    eSYSTEMROOT         'The Windows XP root directory, usually C:\Windows
'    eTEMP               'C:\Users\{Username}\AppData\Local\Temp
'    eUSERNAME           '{username}
'    eUSERPROFILE        'C:\Users\{username}
'    eWINDIR             'C:\Windows
'End Enum

Private Enum EnvironmentVariableName
'WARNING: if any values are added, need to reset corresponding array in ChangeCurrentDir()
'WARNING: Using Option Base 1 will break this!
'http://vlaurie.com/computers2/Articles/environment.htm
'http://vlaurie.com/computers2/Articles/environment-variables-windows-vista-7.htm
    eALLUSERSPROFILE = 1   'C:\ProgramData
    eAPPDATA     'C:\Users\(username}\AppData\Roaming
    eCommonProgramFiles     'C:\Program Files\Common Files
    eCOMPUTERNAME     '{computername}
    eCOMSPEC     'C:\Windows\System32\cmd.exe
    eHOMEDRIVE     'C:
    eHOMEPATH     '\Users\{username}
    eLOCALAPPDATA     'C:\Users\{username}\AppData\Local
    ePATH     'Varies. Includes C:\Windows\System32\;C:\Windows\;C:\Windows\System32\Wbem
    ePATHEXT     '.COM; .EXE; .BAT; .CMD; .VBS; .VBE; .JS ; .WSF; .WSH; .MSC
    eProgramData     'C:\ProgramData
    eProgramFiles     'Directory containing program files, usually C:\Program Files
    eProgramFilesX86   'In 64-bit systems, directory containing 32-bit programs. Usually C:\Program Files (x86)
    ePublic     'C:\Users\Public
    eSYSTEMDRIVE     'The drive containing the Windows XP root directory, usually C:
    eSYSTEMROOT     'The Windows XP root directory, usually C:\Windows
    eTEMP     'C:\Users\{Username}\AppData\Local\Temp
    eUSERNAME     '{username}
    eUSERPROFILE     'C:\Users\{username}
    eWINDIR     'C:\Windows
End Enum

Private Enum PathParseMode
    Path
    FileName
    FileExtension
    FileNameWithoutExtension
End Enum

Private Enum enPriority_Class
    NORMAL_PRIORITY_CLASS = &H20
    IDLE_PRIORITY_CLASS = &H40
    HIGH_PRIORITY_CLASS = &H80
End Enum

'1 based time units
Public Enum ChartTimeUnit '(iTime)
    eUndefined = -1
    eSeconds = 1
    eMinutes = 2
    eHours = 3
    eDays = 4
    eWeeks = 5
End Enum

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'Use instead of SendMessage so that application doesn't hang waiting for response if destination window is unresponsive.
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long

Private Declare Function GetProcessID Lib "kernel32" Alias "GetProcessId" (ByVal hProc As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwprocessid As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetTopWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Private Declare Function GetDlgItem Lib "user32" (ByVal hDlg As Long, ByVal nIDDlgItem As Long) As Long

Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32.dll" (rclsid As GUID) As Long

Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
'Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProc As Long, bWOW64Process As Boolean) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function OpenProcess Lib "kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpAppName As String, _
    ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, _
    ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, lpStartupInfo As STARTUPINFO, _
    lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Const GWL_STYLE = -16

Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_NOSIZE = &H1

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Const SHCONTF_FOLDERS = &H20
Private Const SHCONTF_NONFOLDERS = &H40
Private Const SHCONTF_INCLUDEHIDDEN = &H80

Private Const WM_USER = &H400                        ' 0x0400 'used by applications to define private messages
Private Const WM_CLOSE = &H10
Private Const WM_GETTEXT = &HD

Private Const UNIQUE_NAME = &H0
Private Const MAX_PATH = 260

Private Const STARTF_USESHOWWINDOW = &H1

Private Const SYNCHRONIZE = &H100000
Private Const PROCESS_TERMINATE = &H1

Private Const WS_DISABLED = &H8000000

Private Test As Variant 'For testing purposes.

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' File/Directory/Path Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function DoesFilePathExist(strPath As String, Optional IsDirectory As Boolean = False) As Boolean
'Replacement for Dir() function that can return error #52 "Bad file name or number" when testing for invalid file with a network path.
    On Error GoTo errsub:
    
    Dim FSO As Object 'FileSystemObject
     
    Set FSO = CreateObject("Scripting.Filesystemobject")
    
    If IsDirectory = False Then
        DoesFilePathExist = FSO.FileExists(strPath)
    Else
        DoesFilePathExist = FSO.FolderExists(strPath)
    End If
errsub:
    Set FSO = Nothing
End Function

Private Function ChangeCurrentDir(EnvironmentVariable As EnvironmentVariableName) As Boolean
'Change current directory to environment variable, or system file variable.
    Call ChDir(GetEnvironPath(EnvironmentVariable))
End Function

Private Function GetEnvironPath(EnvironmentVariable As EnvironmentVariableName) As String
    GetEnvironPath = VBA.Environ(Array("ALLUSERSPROFILE", "APPDATA", "CommonProgramFiles", "COMPUTERNAME", "COMSPEC", "HOMEDRIVE", "HOMEPATH", "LOCALAPPDATA", "PATH", "PATHEXT", "ProgramData", "PROGRAMFILES", "ProgramFiles(x86)", "Public", "SYSTEMDRIVE", "SYSTEMROOT", "TEMP", "USERNAME", "USERPROFILE", "WINDIR")(EnvironmentVariable))
End Function

Private Function GetAbsolutePathName(ByVal strPath As String) As String
'http://msdn.microsoft.com/en-us/library/zx1xa64f(v=vs.85).aspx
'Returns a complete and unambiguous path from a provided path specification.
'Resolves "..", *, etc. in file names and returns human readable path.
    On Error GoTo errsub

    Dim FSO As Object

    Set FSO = CreateObject("Scripting.Filesystemobject")
    GetAbsolutePathName = FSO.GetAbsolutePathName(strPath)

errsub:
    Set FSO = Nothing
End Function

Private Function GetTempFileName(Optional strFileExtension As String = "tmp") As String
'When called using UNIQUE_NAME creates unique temp file name.
'WinAPI GetTempPath() can also be used to get temp path instead of VBA.Environ("TEMP").
'Returns full path name to unique file.
    On Error GoTo errsub
    Dim strResult As String

    strResult = VBA.Space(MAX_PATH)
    Call GetTempFileNameA(VBA.Environ("TEMP"), "TMP", UNIQUE_NAME, strResult)
    strResult = Left$(strResult, InStr(strResult, VBA.Chr(0)) - 1)
    If DoesFilePathExist(strResult) = True Then 'File creation ensures unique file name.
        Call Kill(strResult)
        strFileExtension = VBA.Replace(strFileExtension, ".", "")
        GetTempFileName = VBA.Left(strResult, Len(strResult) - 3) & strFileExtension
    End If
errsub:
End Function

Private Function ParsePath(ByVal strPath As String, iMode As PathParseMode) As String
    'Take the path passed in and return the filename, or the path base upon iMode.
    'Path returns with trailing Application.PathSeparator ("\")
    'If file name is passed in to returnPath, empty string is returned.

    On Error GoTo errsub

    Dim FSO As Object
    Set FSO = CreateObject("Scripting.Filesystemobject")

    Select Case iMode
        Case PathParseMode.FileExtension
            ParsePath = FSO.GetExtensionName(strPath) 'File extension name
        Case PathParseMode.FileName
            ParsePath = FSO.GetFileName(strPath) 'File name with extension
        Case PathParseMode.Path
            If InStr(1, strPath, "\") > 0 Then  '"\" exists
                strPath = FSO.GetParentFolderName(strPath) 'File path
                ParsePath = FSO.BuildPath(strPath, "\") 'Add "\"
            End If
        Case PathParseMode.FileNameWithoutExtension
            ParsePath = FSO.GetBaseName(strPath) 'File name without extension
    End Select
errsub:
    Set FSO = Nothing
End Function

Private Function ParsePathLegacy(ByVal strPath As String, iMode As PathParseMode) As String
''Take the path passed in and return the filename, or the path base upon iMode.
''Path returns with trailing Application.PathSeparator ("\")
''Function has weakness in determining file/path names when path contains "."
'
'    strPath = VBA.Trim(strPath)
'
'    If InStr(1, strPath, "\") > 0 Then  '"\" exists
'        If iMode = PathParseMode.returnFileName Then 'File Name
'            If VBA.InStr(1, VBA.Mid(strPath, VBA.InStrRev(strPath, "\") + 1), ".") > 0 Then '"." exists, assumed to be file name.
'                ParsePath = VBA.Mid(strPath, VBA.InStrRev(strPath, "\") + 1)
'            End If
'        ElseIf iMode = PathParseMode.returnPath Then 'File Path
'            If VBA.InStr(VBA.InStrRev(strPath, "\") + 1, strPath, ".") > 0 Then '"." does not exist, assumed to be path name.
'                ParsePath = VBA.Left(strPath, VBA.InStrRev(strPath, "\") - 1)
'            Else
'                ParsePath = strPath
'            End If
'            If VBA.Right(ParsePath, 1) <> "\" Then
'                ParsePath = ParsePath & "\"
'            End If
'        ElseIf iMode = PathParseMode.returnFileExtension Then
'            If VBA.InStr(InStrRev(strPath, "\") + 1, strPath, ".") > 0 Then '"." exists, assumed to be file name.
'                ParsePath = VBA.Mid(strPath, VBA.InStrRev(strPath, ".") + 1)
'            End If
'        End If
'    Else 'Path consists of only file name or path. No included "\"
'        If iMode = PathParseMode.returnFileName Then
'            If VBA.InStr(1, strPath, ".") > 0 Then '"." exists, assumed to be file name.
'                ParsePath = strPath
'            End If
'        ElseIf iMode = PathParseMode.returnPath Then
'            If VBA.InStr(1, strPath, ".") = 0 Then '"." does not exist, assumed to be path name.
'                ParsePath = strPath
'            End If
'            If VBA.Right(ParsePath, 1) <> "\" And VBA.Len(ParsePath) > 0 Then
'                ParsePath = ParsePath & "\"
'            End If
'        ElseIf iMode = PathParseMode.returnFileExtension Then
'            If VBA.InStr(1, strPath, ".") > 0 Then '"." exists, assumed to be file name.
'                ParsePath = VBA.Mid(strPath, VBA.InStrRev(strPath, ".") + 1)
'            End If
'        End If
'    End If
End Function

Private Function CreateFilePath(strFilePath As String) As String
'Only creates one folder level.
    Dim strFolderPath As String

    If strFilePath <> vbNullString Then
        strFolderPath = ParsePath(strFilePath, Path)
        If DoesFilePathExist(strFolderPath, True) = False Then
            Call VBA.MkDir(strFolderPath)
        End If
    End If

    CreateFilePath = strFolderPath
End Function

Private Function MakeDirectoryFSO(strFullFilePath As String) As Boolean
'Untested
''Only creates one folder level.
'    On Error GoTo ErrSub
'
'    Dim FSO As Object
'    Dim strFullPath As String
'
'    Set FSO = CreateObject("Scripting.Filesystemobject")
'
'    strFullPath = ParsePath(strFullFilePath, Path)
'
'    If DoesFilePathExist(strFolderPath, True) = False Then
'        Call FSO.CreateFolder(strFullPath)
'    End If
'    MakeDirectory = True
'
'ErrSub:
'    Set FSO = Nothing
End Function

Private Function MakeDirectory(strFolderPath As String) As Boolean
'Create a subfolder under current file folder to hold generated graphs
    On Error GoTo errsub

    If DoesFilePathExist(strFolderPath, True) = False Then
        Call VBA.MkDir(strFolderPath) 'Create the subfolder
    End If
    MakeDirectory = True
errsub:
End Function

Private Function DeleteDirectory(FileNameFolder As String) As Boolean
'Delete directory and contents.
    Dim FSO As Object

    On Error GoTo errsub

    If DoesFilePathExist(FileNameFolder, True) = True Then
        Set FSO = CreateObject("Scripting.Filesystemobject")
        Call FSO.deletefolder(FileNameFolder)
    End If
    DeleteDirectory = True

errsub:
    Set FSO = Nothing
End Function

Private Function FindFilesCollectionByExtension(ByVal strPath As String, ByVal strFileExtension As String) As Collection
'Returns collection of files in a directory with the given strFileExtension
'Dir() may fail on network path.
    Dim strFile As String
    Dim colResults As New Collection

    strPath = VBA.Trim(strPath)

    If VBA.Right(strPath, 1) <> Application.PathSeparator Then
        strPath = strPath & Application.PathSeparator
    End If
    
    strFile = VBA.Dir(strPath & "*." & strFileExtension) 'Find initial file name
    Do Until (strFile = vbNullString)
        Call colResults.Add(strPath & strFile, strFile)
        strFile = VBA.Dir 'Get additional file name
    Loop
    Set FindFilesCollectionByExtension = colResults
End Function

Private Function PopulateDirectoryList(strCurrentDir As String) As Variant
'Function PopulateDirectoryList(strCurrentDir As String, bFullPath As Boolean) As Variant ', bSort As Boolean) As Variant
'Get unsorted array of folder objects in a directory.
'http://msdn2.microsoft.com/en-us/library/2z9ffy99(VS.85).aspx

    Dim objFSO As Object
    Dim fCurrent As Variant 'Folder
    Dim fSub As Variant 'Folder

    Dim aryFolderList() As Variant 'Folder
    Dim i As Long

    Set objFSO = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
    Set fCurrent = objFSO.GetFolder(strCurrentDir)

    If fCurrent.SubFolders.Count > 0 Then
        ReDim aryFolderList(fCurrent.SubFolders.Count - 1)

        'Iterate through subfolders.
        For Each fSub In fCurrent.SubFolders
            Set aryFolderList(i) = fSub '.path
            i = i + 1
        Next
        PopulateDirectoryList = aryFolderList
'    Else
'        Set PopulateDirectoryList = Nothing
    End If

   'Removed, due to need for alphanumeric sort here.
   'Sort array
'    If bSort = True Then
'        Dim ii As Long
'        Dim str1 As String
'        Dim str2 As String
'
'        For i = 0 To UBound(aryFolderList)
'           For ii = i To UBound(aryFolderList)
'                If VBA.UCase(aryFolderList(ii)) < VBA.UCase(aryFolderList(i)) Then
'                    str1 = aryFolderList(i)
'                    str2 = aryFolderList(ii)
'                    aryFolderList(i) = str2
'                    aryFolderList(ii) = str1
'                End If
'            Next ii
'        Next i
'    End If
End Function

''Private Sub TestFunction()
''    Call EnumerateFolders("c:\")
''End Sub
''
''Private Sub fnFolderItems3FilterVB()
''    Dim objShell As Object 'Shell
''    Dim objFolder As Object 'Folder
''    Dim ssfWINDOWS As Long
''
''    ssfWINDOWS = 36
''    Set objShell = CreateObject("Shell.Application")
'''    Set objShell = New Shell
''    Set objFolder = objShell.Namespace(ssfWINDOWS)
''        If (Not objFolder Is Nothing) Then
''            Dim objFolderItems3 As Object 'FolderItems3
''
''            Set objFolderItems3 = objFolder.Items
''                If (Not objFolderItems3 Is Nothing) Then
''                    Dim SHCONTF_NONFOLDERS As Long
''
''                    SHCONTF_NONFOLDERS = 64
''
''                    Debug.Print objFolderItems3.Count
''                    objFolderItems3.Filter SHCONTF_NONFOLDERS, "*.exe"
''                    Debug.Print objFolderItems3.Count
''                End If
''            Set objFolderItems3 = Nothing
''        End If
''    Set objFolder = Nothing
''    Set objShell = Nothing
''End Sub
''
''Private Sub EnumerateFolders(DirToScan As Variant)
'''Loop through folders using shell object, which is faster than FSO.
'''Untested.
'''http://www.codeguru.com/forum/showthread.php?t=497321&page=7
'''Reference to: shell32.dll
''    Dim FolderCollection As Object
''    Dim SingleFolder As Object
''    Dim clItems As Object 'FolderItems3
''
''    Set clItems = CreateObject("Shell.Application")
''    Set FolderCollection = CreateObject("Shell.Application").Namespace(DirToScan)
'''    Set FolderCollection = ShellObject.Namespace(DirToScan) 'Early bind
''
'''    clItems.Filter SHCONTF_FOLDERS Or SHCONTF_NONFOLDERS Or SHCONTF_INCLUDEHIDDEN, "*"
'''    For Each oFolderItem In clItems
'''        If oFolderItem.IsFileSystem Then
'''            If oFolderItem.IsFolder Then
'''                colFolders.Add oFolderItem.GetFolder
'''                numDirs = numDirs + 1
'''            Else
'''                NumFiles = NumFiles + 1
'''            End If
'''        End If
'''    Next oFolderItem
''End Sub

Private Function LoadFile(strFilePath As String, Optional appStyle As nShowCmd = nShowCmd.SW_SHOWNORMAL) As Long
'Loads an exe or file with the default provider.
'Retuens Main windows handle if exe.
'http://support.microsoft.com/kb/238245
'Code to make process wait until return when calling Shell.
'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=6071&lngWId=1
    On Error GoTo errsub
    Dim PID As Long
    Dim colWindows As Collection
    Dim TimeoutSeconds As Long
    Dim dateCurrentTime As Date

    TimeoutSeconds = 5

    If DoesFilePathExist(strFilePath) = True Then
        If VBA.UCase(ParsePath(strFilePath, FileExtension)) = "EXE" Then 'Launch an exe
            PID = Shell(strFilePath, appStyle) 'Returns Task ID (AKA Process ID) 'Can error if path to call is not found.
            'http://www.programmersheaven.com/mb/VBasic/167592/167592/return-value-of-shell/
            dateCurrentTime = Now()
            Do While colWindows Is Nothing Or DateDiff("s", dateCurrentTime, Now) > TimeoutSeconds
                VBA.Interaction.DoEvents
                Call Sleep(100) 'in milliseconds
                Set colWindows = GetMainHWndCollectionByProcessID(PID, True)
            Loop
            If Not colWindows Is Nothing Then
                LoadFile = colWindows(1)
            End If
        Else 'Open a directory
            'http://msdn.microsoft.com/en-us/library/bb762153(VS.85).aspx
            'Returns non-standard hInstance as Integer, used to filter error code.
            'CoInitializeEx(NULL, COINIT_APARTMENTTHREADED | COINIT_DISABLE_OLE1DDE)
            LoadFile = CLng(ShellExecute(0&, vbNullString, strFilePath, vbNullString, vbNullString, appStyle))
        End If
    End If
errsub:
End Function

Private Function ProcessIDTohWnd(ByVal target_pid As Long) As Long
'Get Hwnd from process ID.
    Dim test_hwnd As Long
    Dim test_pid As Long
    Dim test_thread_id As Long

    'Find the first window
    test_hwnd = GetTopWindow(0&)
'    test_hwnd = FindWindow(ByVal 0&, ByVal 0&)
    Do While test_hwnd <> 0
        'Check if the window isn't a child
        If GetParent(test_hwnd) = 0 Then
            'Get the window's thread
            test_thread_id = GetWindowThreadProcessId(test_hwnd, test_pid)

            If test_pid = target_pid Then
                ProcessIDTohWnd = test_hwnd
                Exit Do
            End If
        End If
        'retrieve the next window
'        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
        test_hwnd = GetNextWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Function

Private Sub WriteDataToFile(strFilePath As String, FileBuffer As String)
    'Write data to file quickly
    'http://www.avdf.com/apr98/art_ot003.html

    'to write the data dimension and fill array A, then ....
    Dim FileNumber As Integer
    FileNumber = FreeFile

    Open strFilePath For Binary As #FileNumber
    Put #FileNumber, , FileBuffer 'writes whole of A to file
    'Put [#]filenumber, [recnumber], varname
    Close #FileNumber
End Sub

Private Sub ReadDataFromFile(strFilePath As String, FileBuffer As String)
    'Read data to file quickly
    'http://www.avdf.com/apr98/art_ot003.html

    'to read it back
'    Dim A(30, 10) As Single
    Dim FileNumber As Integer
    FileNumber = FreeFile

    Open strFilePath For Binary As #FileNumber
    Get #FileNumber, , FileBuffer 'reads whole of A
    Close #FileNumber
End Sub

Private Function GetFileDateModified(strFullFileName As String) As Date
On Error GoTo errsub
    With CreateObject("Scripting.FileSystemObject").GetFile(strFullFileName)
        GetFileDateModified = .DateLastModified
        '.DateLastAccessed
        '.DateCreated
        '.Path
        '.ShortPath
        '.ShortName
    End With
errsub:
End Function

Private Sub cmdWriteValues_Click()
'http://www.vb-helper.com/howto_read_write_binary_file.html
'    Dim file_name As String
'    Dim file_length As Long
'    Dim fnum As Integer
'    Dim bytes() As Byte
'    Dim txt As String
'    Dim i As Integer
'    Dim values As Variant
'    Dim num_values As Integer
'
'    ' Build the values array.
'    values = Split(txtValues.Text, vbCrLf)
'    For i = 0 To UBound(values)
'        If Len(Trim$(values(i))) > 0 Then
'            num_values = num_values + 1
'            ReDim Preserve bytes(1 To num_values)
'            bytes(num_values) = values(i)
'        End If
'    Next i
'
'    ' Delete any existing file.
'    file_name = txtFile.Text
'    On Error Resume Next
'    Kill file_name
'    On Error GoTo 0
'
'    ' Save the file.
'    fnum = FreeFile
'    Open file_name For Binary As #fnum
'    Put #fnum, 1, bytes
'    Close fnum
'
'    ' Clear the results.
'    txtValues.Text = ""
End Sub

Private Function ReadBinaryFile(strFilePath As String) As Byte()
'Read binary file and return byte array
'http://www.vb-helper.com/howto_read_write_binary_file.html
    Dim intFileNum As Integer
    Dim bytes() As Byte

    On Error GoTo errsub
    intFileNum = FreeFile

    If DoesFilePathExist(strFilePath) = True Then
        Open strFilePath For Binary Access Read As #intFileNum
        ReDim bytes(1 To FileLen(strFilePath))
        Get #intFileNum, 1, bytes
        ReadBinaryFile = bytes()
    End If

errsub:
    Close intFileNum
End Function

Private Function ByteArrayToString(arryByte() As Byte) As String
    'All VB string are Unicode
    ByteArrayToString = VBA.StrConv(arryByte(), vbUnicode) 'To convert from Unicode
End Function

Private Function ExtractOLEObjectToFolder(obj As OLEObject, strDestination As String) As Boolean
'Used to extract embedded OLE object (in office applications) to files in a directory utilizing clipboard object.
'Could also create temp folder with CreateFilePath() and delete with DeleteDirectory().
'http://www.excelforum.com/excel-programming/387216-pasting-embedded-ole-objects-to-specific-filename.html
    On Error GoTo errsub

    Dim oClipData As Object 'DataObject
'    Dim oApp As Object
'    Dim oFolder As Object
    Dim FileNameFolder As Variant

    FileNameFolder = ParsePath(strDestination, Path)

'    Set oApp = CreateObject("Shell.Application")
'    Set oFolder = oApp.Namespace(FileNameFolder)

    Set oClipData = GetClipboard 'Backup clipboard contents.
    Call obj.Copy
'    Call oFolder.Self.InvokeVerb("Paste")
    Call CreateObject("Shell.Application").Namespace(FileNameFolder).Self.InvokeVerb("Paste")

    Call SetClipboard(oClipData) 'Restore clipboard contents.

    Exit Function

errsub:
    Debug.Print "Error in ExtractOLEObjectToFolder()"
    Call SetClipboard(oClipData) 'Restore clipboard contents.
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Control/Object/Collection Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function RegisterInROT(hWnd As Long)
    'Use the SendMessage API function to enter Application into Running Object Table.
    'If multiple instances are running, only the first launched is entered, so enter in ROT.
    Call SendMessage(hWnd, WM_USER + 18, 0, 0)
End Function

Private Function GetControlValue(ControlName As Variant) As Variant
'    Dim RetVal As Variant
'    Dim strControlName As String
'
''    Application.Volatile True
'
'    If IsObject(ControlName) Then
'        If TypeOf ControlName Is Excel.Range Then
'            strControlName = ControlName.Value
'        Else
'            RetVal = "#ERROR"
'        End If
'    Else
'        strControlName = ControlName
'    End If
'
'    If RetVal <> "#ERROR" Then
'        GetControlValue = Sheet2.OLEObjects(strControlName).Object.Value
'    Else
'        GetControlValue = RetVal
'    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' General Utility Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function CollectionContainsItem(ColTest As Collection, varTemp As Variant) As Boolean
'Check to see if a collection contains an item.
On Error GoTo errsub
    If TypeOf ColTest.Item(varTemp) Is Object  Or Not TypeOf ColTest.Item(varTemp) Is Object  Then
    End If
errsub:
    CollectionContainsItem = Not (Err.Number = 5 Or Err.Number = 9)
End Function

Private Function ArrayToCollection(yArray As Variant) As Collection
'Pass in yArray with Array()

    Dim cReturn As Collection
    Dim i As Long

    If UBound(yArray) - LBound(yArray) >= 0 Then
        Set cReturn = New Collection

        For i = LBound(yArray) To UBound(yArray)
            Call cReturn.Add(yArray(i))
        Next

        Set ArrayToCollection = cReturn
    End If
End Function

Private Function CollectionToArray(col As Collection) As Variant
    Dim lCount As Long
    Dim vArray As Variant
    Dim vItem As Variant

    ReDim vArray(col.Count - 1) As Variant

    For Each vItem In col
        vArray(lCount) = vItem
        lCount = lCount + 1
    Next

    CollectionToArray = vArray
End Function

Private Function RemoveCollectionDuplicates(ByVal cOriginal As Collection, cNew As Collection) As Collection
'Remove duplicate New collection items from Original collection
    Dim cItemOrig As Variant
    Dim cItemNew As Variant
    Dim i As Long

    If Not cOriginal Is Nothing And Not cNew Is Nothing Then
'        Set RemoveCollectionDuplicates = cItemNew
        For Each cItemNew In cNew
            i = 1
            For Each cItemOrig In cOriginal
                If StrComp(cItemOrig, cItemNew) = 0 Then 'Same
                    Call cOriginal.Remove(i)
                    Exit For
                End If
                i = i + 1
            Next
        Next

        If cOriginal.Count > 0 Then
            Set RemoveCollectionDuplicates = cOriginal
        End If
    ElseIf cOriginal Is Nothing And Not cNew Is Nothing Then
        If cNew.Count > 0 Then
            Set RemoveCollectionDuplicates = cNew
        End If
    End If
End Function

Private Function IsElementInArray(arryTest, Element, Optional ByRef Index = 0) As Boolean
'Test if value is in an Array.
'Returns (1 based) Index
On Error GoTo errsub 'Array could be empty
    Dim arryElement
    For Each arryElement In arryTest
        Index = Index + 1
        If Element = arryElement Then
            IsElementInArray = True
            Exit For
        End If
    Next
errsub:
End Function

Private Sub ClearClipboard() '(lHwnd As Long)
'Completely clear the contents of the clipboard using only Windows API calls.
'http://www.cpearson.com/excel/clipboar.htm for clipboard function.
    Dim lResult As Long         'lResult of 0 is failure
    lResult = OpenClipboard(0&) '0&) '0& = handle to desktop passed in as a Long.
    lResult = EmptyClipboard()
    lResult = CloseClipboard()
End Sub

Private Function GetClipboard() As Object 'DataObject
    On Error GoTo errsub

    Dim oData As Object 'DataObject
    Set oData = GetDataObject() 'DataObject
    
    Call oData.GetFromClipboard
    Set GetClipboard = oData
errsub:
End Function

Private Function SetClipboard(oData As Object) As Boolean 'DataObject
    On Error GoTo errsub

    oData.PutInClipboard
    SetClipboard = True

errsub:
End Function

Private Function GetDataObject() As Object
    'Create new object by class GUID
    'This is ok when you use a UserForm referenced with "Microsoft Forms 2.0 Object Library" - FM20.dll
    'Dim oDataObject As MSForms.DataObject
    'Set oDataObject = New MSForms.DataObject
    
    Set GetDataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'Microsoft.Vbe.Interop.Forms.DataObjectClass
'    Set GetDataObject = CreateObject("Application.Forms")
End Function

'Sub ClearClipboard()
'Required reference to Windows/System32/FM20.dll
''    Dim MyDataObj As DataObject
''    Set MyDataObj = New DataObject
'    MyDataObj.SetText ""
'    MyDataObj.PutInClipboard
'End Sub

'Function GetOffClipboard() As Variant
''http://www.cpearson.com/excel/clipboar.htm
''    Dim MyDataObj As DataObject
''    Set MyDataObj = New DataObject
'    MyDataObj.GetFromClipboard
'    GetOffClipboard = MyDataObj.GetText()
'End Function

'Function GetClipboard(strText As String, Optional vFormat As Variant = Empty) As Variant
''Required reference to Windows/System32/FM20.dll
''http://www.cpearson.com/excel/clipboar.htm
''References - FM20.dll - "Microsoft Forms 2.0 Object Library" for DataObject
''    'Dim oData As DataObject 'object to use the clipboard
''    'Set oData = New DataObject 'object to use the clipboard
'    Dim oData As Object
''    Set oData = CreateObject("MSForms.DataObject")
'    Set oData = CreateObject("Application.Forms")
''
'    Call oData.GetText(strText, vFormat)
'    oData.PutInClipboard                    'take in the clipboard to empty it
'    Set oData = Nothing
'End Function

'Sub PutOnClipboard(Obj As Variant)
''Required reference to Windows/System32/FM20.dll
''http://www.cpearson.com/excel/clipboar.htm
'    Dim MyDataObj As DataObject
'    Set MyDataObj = New DataObject
'    MyDataObj.SetText Format(Obj)
'    MyDataObj.PutInClipboard
'End Sub

'Sub SetClipboard(strText As String, Optional vFormat As Variant = Empty)
'http://www.cpearson.com/excel/clipboar.htm
''References - FM20.dll - "Microsoft Forms 2.0 Object Library" for DataObject
'    'Dim oData As DataObject 'object to use the clipboard
'    'Set oData = New DataObject 'object to use the clipboard
'    Dim oData As Variant 'Object
'    Set oData = CreateObject("MSForms.DataObject")
'
'    Call oData.SetText(strText, vFormat)    'Set to Empty to clear clipboard contents
'    oData.PutInClipboard                    'take in the clipboard to empty it
'    Set oData = Nothing
'End Sub

Private Sub GetRGB(RGB As Long, ByRef Red As Integer, ByRef Green As Integer, ByRef Blue As Integer)
'Method to break out the individual color components from a color value.
'http://www.cpearson.com/excel/colors.htm
    Red = RGB And 255
    Green = RGB \ 256 And 255
    Blue = RGB \ 256 ^ 2 And 255
End Sub

Private Sub pRedraw(mHwnd As Long)
    ' Redraw window
    Const swpFlags As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE
    Call SetWindowPos(mHwnd, 0, 0, 0, 0, 0, swpFlags)
End Sub

Private Function IsArrayEmpty(checkArray As Variant, Optional lDimension As Long = 1) As Boolean
'Used to tell is array is empty. Uninitialized arrays or arrays that are cleared with Erase() return true.
    On Error GoTo emptyError 'Sometimes errors when testing if = -1 if empty.

    If -1 = UBound(checkArray, lDimension) Then
        GoTo emptyError
    End If

    Exit Function
emptyError:
    Err.Clear
    IsArrayEmpty = True

'End Function
'
'Private Function IsArrayEmpty(checkArray As Variant) As Boolean
''Used to tell is array is empty.
''Uninitialized arrays or arrays that are cleared with Erase().
'    Dim lngTmp As Long
'    On Error GoTo emptyError
'
'    'Here is where it happens.
'    'If you haven't used Redim on your array
'    'Ubound will return an error
'    lngTmp = UBound(checkArray)
'    If lngTmp = -1 Then GoTo emptyError 'Returns -1 if empty.
'    IsArrayEmpty = False
'    Exit Function
'emptyError:
'    Err.Clear 'Clear out error code
'    On Error GoTo 0 'Turn off error checking
'    IsArrayEmpty = True
End Function

Private Function GetArrayMin(vArray As Variant) As Variant
    Dim i As Long
    Dim vMin As Variant

    vMin = vArray(LBound(vArray))
    For i = LBound(vArray) To UBound(vArray)
        If vArray(i) < vMin Then
            vMin = vArray(i)
        End If
    Next
    GetArrayMin = vMin
End Function

Private Function GetArrayMax(vArray As Variant) As Variant
    Dim i As Long
    Dim vMax As Variant

    vMax = vArray(LBound(vArray))
    For i = LBound(vArray) To UBound(vArray)
        If vArray(i) > vMax Then
            vMax = vArray(i)
        End If
    Next
    GetArrayMax = vMax
End Function

Private Sub SortComboBoxList(ByRef cmbSource As Object) 'ComboBox
'   Sorts a ComboBox using bubble sort algorithm
    Dim First As Integer, Last As Integer
    Dim i As Integer, j As Integer
    Dim Temp(0 To 1) As String
    Dim List As Variant

    First = LBound(cmbSource.List)
    Last = UBound(cmbSource.List)
    For i = First To Last - 1
        For j = i + 1 To Last
            If VBA.UCase(cmbSource.List(i, 0)) > VBA.UCase(cmbSource.List(j, 0)) Then
                Temp(0) = cmbSource.List(j, 0)
                Temp(1) = cmbSource.List(j, 1)

                cmbSource.List(j, 0) = cmbSource.List(i, 0)
                cmbSource.List(j, 1) = cmbSource.List(i, 1)

                cmbSource.List(i, 0) = Temp(0)
                cmbSource.List(i, 1) = Temp(1)
            End If
        Next j
    Next i
End Sub

Private Sub BubbleSort(List() As Integer)
'   Sorts an array using bubble sort algorithm

    Dim First As Integer, Last As Integer
    Dim i As Integer, j As Integer
    Dim Temp As Integer

    First = LBound(List)
    Last = UBound(List)
    For i = First To Last - 1
        For j = i + 1 To Last
            If List(i) > List(j) Then
                Temp = List(j)
                List(j) = List(i)
                List(i) = Temp
            End If
        Next j
    Next i
End Sub

Private Sub QuickSort(List() As Integer)
'   Sorts an array using Quick Sort algorithm
'   Adapted from "Visual Basic Developers Guide"
'   By D.F. Scott

    Dim i As Integer, j As Integer, b As Integer, k As Integer
    Dim l As Integer, t As Integer, r As Integer, d As Integer
    Dim comp As Integer, swic As Integer
    Dim oldx1 As Integer, oldy1 As Integer, oldx2 As Integer, oldy2 As Integer
    Dim newx1 As Integer, newy1 As Integer, newx2 As Integer, newy2 As Integer

    Dim p(1 To 100) As Integer
    Dim w(1 To 100) As Integer

    k = 1
    p(k) = LBound(List)
    w(k) = UBound(List)
    l = 1
    d = 1
    r = UBound(List)
    Do
toploop:
        If r - l < 9 Then GoTo bubsort
        i = l
        j = r
        While j > i
           comp = comp + 1
           If List(i) > List(j) Then
               swic = swic + 1
               t = List(j)
               oldx1 = List(j)
               oldy1 = j
               List(j) = List(i)
               oldx2 = List(i)
               oldy2 = i
               newx1 = List(j)
               newy1 = j
               List(i) = t
               newx2 = List(i)
               newy2 = i
               d = -d
           End If
           If d = -1 Then
               j = j - 1
                Else
                    i = i + 1
           End If
       Wend
           j = j + 1
           k = k + 1
            If i - l < r - j Then
                p(k) = j
                w(k) = r
                r = i
                Else
                    p(k) = l
                    w(k) = i
                    l = j
            End If
            d = -d
            GoTo toploop
bubsort:
    If r - l > 0 Then
        For i = l To r
            b = i
            For j = b + 1 To r
                comp = comp + 1
                If List(j) <= List(b) Then b = j
            Next j
            If i <> b Then
                swic = swic + 1
                t = List(b)
                oldx1 = List(b)
                oldy1 = b
                List(b) = List(i)
                oldx2 = List(i)
                oldy2 = i
                newx1 = List(b)
                newy1 = b
                List(i) = t
                newx2 = List(i)
                newy2 = i
            End If
        Next i
    End If
    l = p(k)
    r = w(k)
    k = k - 1
    Loop Until k = 0
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Unsorted
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function FormatVBA()
'Documentation for VB Format() function.
'See documentation for formatting options, however there are standard formatting capabilities that are not listed in help.
'These are listed here:
'http://www.vbtutor.net/vb6/lesson12.html

'General Number, Fixed, Standard, Currency, Percent
End Function

Private Function FormatStringForProModel(ByVal strData As String) As String
'In Excel need to set Application.Displayerrors = false so that message doesn't show when no strings are found during VBA.Replace() operation.
    Dim i As Integer
    Dim vArray As Variant

    vArray = Array("+", "-", "*", "/", ",", ":", ";", "(", ")", "[", "]", "{", "}", """, """, "<", ">", "=", "\", "\", "'", "!", "@", "#", "$", "%", "^", "&", "|", "?", "`", "~", ".", vbNullString, vbLf, vbCr, vbTab, VBA.Space(1))

    For i = 0 To UBound(vArray)
        strData = VBA.Replace(strData, CStr(vArray(i)), "_")
    Next

    'First character can't be number
    If VBA.Left(strData, 1) Like "[0-9]" Then
        strData = "_" & VBA.Mid(strData, 2)
    End If

    'String length limit
    If Len(strData) > 74 Then
        strData = VBA.Left(strData, 74)
    End If

    FormatStringForProModel = strData
End Function

'Example to show how to use callback.  Goes with EnumWindowProc()
Private Function GetApplicationHwnd(strTitleString As String, bPartial As Boolean) As Long
'    If bPartial = True Then
'        Call EnumWindows(AddressOf EnumWindowProc, &H0)
'        GetApplicationHwnd = lAppHwnd
'    End If
End Function

'Example for to show how to use callback.  Goes with GetApplicationHwnd()
'Callback must be in module.
Private Function EnumWindowProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
'    Dim sTitle As String
'    Dim sClass As String
'
'    'To continue enumeration, return True
'    'To stop enumeration return False (0).
'    'When 1 is returned, enumeration continues
'    'until there are no more windows left.
'    EnumWindowProc = 1
'
'    'eliminate windows that are not top-level and not visible.
'    If GetParent(hwnd) = 0& And IsWindowVisible(hwnd) Then
'        'get the window title
'        If InStr(1, GetWindowTitle(hwnd), strAppNamePiece) > 0 Then
'            lAppHwnd = hwnd
'            EnumWindowProc = False '0 - stop enumeration return False (0).
'        End If
'    End If
End Function

Private Function GetWindowTitle(ByVal hWnd As Long) As String
'Gets text of window, not necessarily caption.  If you want caption use: GetWindowHandleCollectionByName()
    Dim nSize As Long
    Dim sTitle As String
    Dim retval As Long

    sTitle = VBA.Space(MAX_PATH)
    retval = GetWindowText(hWnd, sTitle, Len(sTitle))
'    retVal = SendMessageTimeout(tempHwnd, WM_GETTEXT, Len(sWindowText), sWindowText, 0&, 15&, 0) + 1 'Return after time.  Don't get stuck waiting for unresponsive window.
    GetWindowTitle = TrimNull(sTitle)
End Function

Private Function GetMainHWndCollectionByName(strName As String, Optional CheckIfVisible As Boolean = False, Optional lProcessID As Long = 0) As Collection
'http://support.microsoft.com/kb/242308
'Looks at visible windows to try and find top most application window handles matching passed in text.
'Intended to find windows that match caption, however WM_GETTEXT can return other text from other window types.
'Optional pass in ProcessID to narrow down to correct instance of application window.

    Dim tempHwnd As Long
    Dim colHwnds As Collection
    Dim sWindowText As String
    Dim retval As String

    Set colHwnds = New Collection

   ' Grab the first window handle that Windows finds (by passing in vbNullString, vbNullString)
    tempHwnd = GetTopWindow(0) 'FindWindow(vbNullString, vbNullString)

    Do Until tempHwnd = 0   ' Loop until you find a match or there are no more window handles:
        If GetParent(tempHwnd) = 0 And GetWindow(tempHwnd, GW_CHILD) <> 0 Then 'Check if no parent for this window and has children.
            sWindowText = GetWindowTitle(tempHwnd)
            If sWindowText <> vbNullString Then
                If sWindowText Like strName & "*" Then
                    If CheckIfVisible = True And IsWindowVisible(tempHwnd) > 0 Then
                        If lProcessID > 0 And lProcessID = GetProcessIDFromHWnd(tempHwnd) Then
                            colHwnds.Add tempHwnd
                        ElseIf lProcessID = 0 Then
                            colHwnds.Add tempHwnd
                        End If
                    ElseIf CheckIfVisible = False Then
                        If lProcessID > 0 And lProcessID = GetProcessIDFromHWnd(tempHwnd) Then
                            colHwnds.Add tempHwnd
                        ElseIf lProcessID = 0 Then
                            colHwnds.Add tempHwnd
                        End If
                    End If
                End If
            End If
        End If
      ' Get the next window handle
        tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT) ' Same as GetNextWindow()
    Loop

    If colHwnds.Count > 0 Then
        Set GetMainHWndCollectionByName = colHwnds
    End If
End Function

Private Function GetMainHWndCollectionByProcessID(lProcessID As Long, Optional CheckIfVisible As Boolean = False) As Collection
'Looks at visible windows to try and find top most application window handles matching passed in text.
'ProcessID can come from "Windows Script Host Object Model"

    Dim tempHwnd As Long
    Dim colHwnds As Collection

    Set colHwnds = New Collection

   ' Grab the first window handle that Windows finds (by passing in vbNullString, vbNullString)
    tempHwnd = GetTopWindow(0) 'FindWindow(vbNullString, vbNullString)

    Do Until tempHwnd = 0   ' Loop until you find a match or there are no more window handles:
        If GetWindowTitle(tempHwnd) <> vbNullString Then
            If GetParent(tempHwnd) = 0 And GetWindow(tempHwnd, GW_CHILD) <> 0 Then 'Check if no parent for this window and has children.
                If lProcessID > 0 And lProcessID = GetProcessIDFromHWnd(tempHwnd) Then
                    If CheckIfVisible = True And IsWindowVisible(tempHwnd) > 0 Then
                        colHwnds.Add tempHwnd
                    ElseIf CheckIfVisible = False Then
                        colHwnds.Add tempHwnd
                    End If
                ElseIf lProcessID = 0 Then
                    If CheckIfVisible = True And IsWindowVisible(tempHwnd) > 0 Then
                        colHwnds.Add tempHwnd
                    ElseIf CheckIfVisible = False Then
                        colHwnds.Add tempHwnd
                    End If
                End If
            End If
        End If
      ' Get the next window handle
        tempHwnd = GetWindow(tempHwnd, GW_HWNDNEXT) ' Same as GetNextWindow()
    Loop

    If colHwnds.Count > 0 Then
        Set GetMainHWndCollectionByProcessID = colHwnds
    End If
End Function

Private Function GetProcessIDFunc(hProcess As Long)
    'Retrieves the process identifier of the specified process.
    GetProcessIDFunc = GetProcessID(hProcess)
End Function

Private Function GetProcessIDFromHWnd(ByVal hWnd As Long) As Long
'http://support.microsoft.com/kb/242308
   Dim idProc As Long
   Dim lThreadID As Long

   lThreadID = GetWindowThreadProcessId(hWnd, idProc) ' Get PID for this HWnd
   GetProcessIDFromHWnd = idProc  ' Return PID
End Function

Private Function GetClassNameFromHWnd(hWnd As Long) As String
'Get class name for passed in window handle.
    Dim lpClassName As String

    lpClassName = VBA.Space(MAX_PATH)
    Call GetClassName(hWnd, lpClassName, Len(lpClassName))
    GetClassNameFromHWnd = VBA.Left(lpClassName, InStr(lpClassName, VBA.Chr(0)) - 1)
End Function

Private Function TrimNull(strToTrim As String) As String
'Trim string contents following null character.
     Dim pos As Long

     TrimNull = strToTrim
     pos = InStr(strToTrim, vbNullChar) 'VBA.Chr(0))
     If pos Then
        TrimNull = VBA.Left(strToTrim, pos - 1)
     End If
End Function

Private Function GetDayOfMonth(DateSource As Date, DayOfMonth As eDayOfMonth) As Date
    Select Case DayOfMonth
        Case eDayOfMonth.FirstDayOfMonth
            GetDayOfMonth = DateSerial(Year(DateSource), Month(DateSource), 1)
        Case eDayOfMonth.LastDayOfMonth
            GetDayOfMonth = DateSerial(Year(DateSource), Month(DateSource) + 1, 1) - 1
    End Select
End Function

Private Function GetNumberOfArrayDimensions(arr As Variant) As Long
    'http://www.cpearson.com/Excel/VBAArrays.htm
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' NumberOfArrayDimensions
    ' This function returns the number of dimensions of an array. An unallocated dynamic array
    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    On Error Resume Next
    Dim Ndx As Integer
    Dim Res As Integer

    ' Loop, increasing the dimension index Ndx, until an error occurs.
    ' An error will occur when Ndx exceeds the number of dimension
    ' in the array. Return Ndx - 1.
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until Err.Number <> 0

    GetNumberOfArrayDimensions = Ndx - 1
    Err.Clear
End Function

Private Function TransposeArray(vArray As Variant) As Variant
'Custom Function to Transpose vArray
'Use instead of Excel's WorksheetFunction.Transpose() due to limitations with call listed at link.
'http://support.microsoft.com/kb/246335
    Dim X As Long
    Dim Y As Long
    Dim Xupper As Long
    Dim Yupper As Long
    Dim XLower As Long
    Dim YLower As Long
    Dim TempArray As Variant

    Yupper = UBound(vArray, 1)
    YLower = LBound(vArray, 1)
    Select Case NumberOfArrayDimensions(vArray)
        Case 1
            ReDim TempArray(0 To 0, YLower To Yupper)
            For Y = YLower To Yupper
                TempArray(0, Y) = vArray(Y)
            Next
        Case 2
            Xupper = UBound(vArray, 2)
            XLower = LBound(vArray, 2)
            ReDim TempArray(XLower To Xupper, YLower To Yupper)
            For X = XLower To Xupper
                For Y = YLower To Yupper
                    TempArray(X, Y) = vArray(Y, X)
                Next
            Next
    End Select

    TransposeArray = TempArray
End Function

Private Function NumberOfArrayDimensions(vArray As Variant) As Long
'Returns the number of dimensions in an array
'http://support.microsoft.com/kb/152288
    On Error GoTo FinalDimension

    Dim DimNum As Long
    Dim ErrorCheck As Long

    'Visual Basic for Applications arrays can have up to 60000 dimensions.
    For DimNum = 1 To 60000
        ErrorCheck = LBound(vArray, DimNum)
    Next

    Exit Function

FinalDimension:
    NumberOfArrayDimensions = DimNum - 1
End Function

Private Function GetGUIDString() As String
    'Generates a new GUID, returning it in canonical (string) format
    'http://www.mrexcel.com/tip078.shtml
    'http://www.cpearson.com/Excel/CreateGUID.aspx
    
    Dim stGuid As String
    Dim rclsid As GUID
    Dim rc As Long
    
    If CoCreateGuid(rclsid) = 0 Then
        stGuid = VBA.String(40, vbNullChar) '39 chars for the GUID plus room for the Null char.
        rc = StringFromGUID2(rclsid, StrPtr(stGuid), VBA.Len(stGuid) - 1)
        GetGUIDString = VBA.Left(stGuid, rc - 1)
    End If
End Function

Private Function StGuidGen() As String
    'http://www.mrexcel.com/tip078.shtml
    'http://www.cpearson.com/Excel/CreateGUID.aspx
    'Generates a new GUID, returning it in canonical (string) format

    Dim rclsid As GUID

    If CoCreateGuid(rclsid) = 0 Then
        StGuidGen = StGuidFromGuid(rclsid)
    End If
End Function

Private Function StGuidFromGuid(rclsid As GUID) As String
    'http://www.mrexcel.com/tip078.shtml
    'http://www.cpearson.com/Excel/CreateGUID.aspx
    'Converts a binary GUID to a canonical (string) GUID.

    Dim rc As Long
    Dim stGuid As String

    ' 39 chars  for the GUID plus room for the Null char
    stGuid = String$(40, vbNullChar)
    rc = StringFromGUID2(rclsid, StrPtr(stGuid), VBA.Len(stGuid) - 1)
    StGuidFromGuid = VBA.Left(stGuid, rc - 1)
End Function

Public Function IsOS64bit() As Boolean
    'CHeck if operating system is 64 bit.
    'Can get other OS information with function that populate OSVERSIONINFO struct
    Dim bWOW64Process As Boolean

    'Check to see if IsWow64Process function exists
    If GetProcAddress(GetModuleHandle("kernel32"), "IsWow64Process") > 0 Then ' IsWow64Process function exists
        'Now use the function to determine if we are running under Wow64
        Call IsWow64Process(GetCurrentProcess(), bWOW64Process)
        IsOS64bit = bWOW64Process
    End If
End Function

Private Function PopulateComboBoxWithRecordset(cmbObj As Object, oRecordSet As Object, strFirstField As String, strSecondField As String, Optional strDefaultSession As String)
    'cmbObj is of type ComboBox
    'Populate ComboBox with specified fields of passed in RecordSet
    'List is created in alphabetical order.
    Dim i As Long
    Dim iMatch As Long

    With cmbObj
        .Clear
        Do While Not oRecordSet.EOF
            'Add to ComboBox in sorted order.
            If UBound(.List) > -1 Then 'Not first item
                For i = LBound(.List) To UBound(.List)
                    If VBA.UCase(.List(i, 0)) >= VBA.UCase(oRecordSet(strFirstField)) Then
                        Exit For
                    End If
                Next i
            End If
            Call .AddItem(oRecordSet(strFirstField), i) 'Normal Name
            .List(i, 1) = oRecordSet(strSecondField) 'ID

            If strDefaultSession = oRecordSet(strFirstField) Then
                iMatch = i
            End If

            oRecordSet.MoveNext
        Loop

        If .ListCount > 1 Then 'Only show if more than one to choose from.  Auto select one if only one.(there must be at least one in Excel workbook)
            '.ListIndex() fires update event of object cmbObj().
            If iMatch <> 0 Then
                .ListIndex = iMatch
            Else
                .ListIndex = 0 'set to first item in list.
            End If
        End If

        PopulateComboBoxWithRecordset = True
    End With
End Function

Private Function CreateProcessEx(ByVal App As String, ByVal WorkDir As String, Optional dwMilliseconds As Long = 1000, Optional ByVal start_size As nShowCmd = nShowCmd.SW_SHOWNORMAL, Optional ByVal Priority_Class As enPriority_Class = enPriority_Class.NORMAL_PRIORITY_CLASS) As PROCESS_INFORMATION
    'http://support.microsoft.com/kb/129796

    Dim StartInfo As STARTUPINFO
    Dim ProcessInfo As PROCESS_INFORMATION

    StartInfo.cb = Len(StartInfo)
    StartInfo.dwFlags = STARTF_USESHOWWINDOW 'Set the flags
    StartInfo.wShowWindow = start_size 'Set the window's startup position

    'Start the program
'    If CreateProcess(vbNullString, sCommandLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, VBA.Environ("TEMP"), Start, ProcessInfo) <> 0 Then
    If CreateProcess(vbNullString, App, 0&, 0&, 1&, Priority_Class, 0&, vbNullString, StartInfo, ProcessInfo) Then
        'Wait
        Call WaitForSingleObject(ProcessInfo.hProcess, dwMilliseconds)
        Call CloseHandle(ProcessInfo.hThread)
        Call CloseHandle(ProcessInfo.hProcess)
        CreateProcessEx = ProcessInfo
    End If
End Function

Private Function TerminateProcessByID(dwProcessId As Long, Optional bWaitForReturn As Boolean = False, Optional WaitTimeoutMilliseconds As Long = 1000) As Long
    'Send WM_Close to all top level windows in process.
    'Note that WaitTimeoutMilliseconds can add up for many windows.
    'Modified from http://support.microsoft.com/kb/178893
    Dim lhwnd As Variant
    Dim hProc As Long

    hProc = OpenProcess(SYNCHRONIZE Or PROCESS_TERMINATE, False, dwProcessId)
    If hProc > 0 Then
        For Each lhwnd In GetMainHWndCollectionByProcessID(dwProcessId, True)
            If bWaitForReturn = True Then
                TerminateProcessByID = SendMessageTimeout(lhwnd, WM_CLOSE, 0&, 0&, 0&, WaitTimeoutMilliseconds, 0)
'                    lRet = SendMessage(TargetAppHwnd, WM_CLOSE, 0&, 0&) 'Wait for return
            Else
                TerminateProcessByID = PostMessage(lhwnd, WM_CLOSE, 0&, 0&) 'Don't wait for return
            End If
            VBA.Interaction.DoEvents 'Causes processing of PostMessage
        Next

        Call WaitForSingleObject(dwProcessId, WaitTimeoutMilliseconds)
        TerminateProcessByID = TerminateProcess(hProc, 0&)
    End If
End Function

Private Function CloseTaskByhWwnd(TargetAppHwnd As Long, Optional bWaitForReturn As Boolean = False, Optional WaitTimeoutMilliseconds As Long = 1000) As Long
'    Debug.Assert False 'If not visible need to force close!
    If TargetAppHwnd <> 0 Then
'        If IsWindow(TargetAppHwnd) = True Then
            If Not (GetWindowLong(TargetAppHwnd, GWL_STYLE) And WS_DISABLED) Then 'Window not disabled.
'                Debug.Print TargetAppHwnd
                If bWaitForReturn = True Then
                    CloseTaskByhWwnd = SendMessageTimeout(TargetAppHwnd, WM_CLOSE, 0&, 0&, 0&, WaitTimeoutMilliseconds, 0)
'                    lRet = SendMessage(TargetAppHwnd, WM_CLOSE, 0&, 0&) 'Wait for return
                Else
                    CloseTaskByhWwnd = PostMessage(TargetAppHwnd, WM_CLOSE, 0&, 0&) 'Don't wait for return
                End If
                VBA.Interaction.DoEvents 'Causes processing of PostMessage
            End If
'        End If
    End If
End Function

Private Function RemoveDuplicatesFrom1DArray(ByVal vArray As Variant) As Double()
    Dim i As Long
    Dim colUnique As New Collection

    For i = 1 To UBound(vArray, 1)
        On Error Resume Next
        colUnique.Add vArray(i), CStr(vArray(i)) 'Will error if duplicate, which removes duplicates.
    Next
    On Error GoTo 0

    ReDim vArray(1 To colUnique.Count) As Double
    For i = 1 To colUnique.Count
        vArray(i) = colUnique(i)
    Next

    RemoveDuplicatesFrom1DArray = vArray
End Function

Private Function LongToUnsigned(ByVal Value As Long) As Double
    If Value < 0 Then
        LongToUnsigned = Value + 4294967296# 'offset 4b
    Else
        LongToUnsigned = Value
    End If
End Function

Private Function GetTimeIntervalLetter(unit As ChartTimeUnit) As String
'Returns time letter specification to use in functions like DataAdd()
    Select Case unit
        Case eSeconds
            GetTimeIntervalLetter = "s"
        Case eMinutes
            GetTimeIntervalLetter = "n"
        Case eHours
            GetTimeIntervalLetter = "h"
        Case eDays
            GetTimeIntervalLetter = "d"
        Case eWeeks
            GetTimeIntervalLetter = "ww"
        Case Else 'eUndefined
            Debug.Print "Error"
    End Select
End Function

Private Sub DownloadFile()
'Untested
'http://officevbavsto-en.blogspot.com.br/2012/09/answer-how-to-download-files-from.html?goback=%2Egde_82527_member_164381729
'May need to call CoInitializeEx() first before calling URLDownloadToFile()
    Dim ie As Object 'InternetExplorer
    Dim objLink As Object
    Dim URL
    
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Navigate "https://www.bom.mu/?ID=80276"
    ie.Visible = True
   
    While ie.Busy
    Wend

    For Each objLink In ie.Document.Links
        If objLink.href Like "*/balance_of_payments/Review*_USD.pdf" Then
            URL = objLink.href
            Exit For
        End If
    Next

    ' Note that this file is 2M, so you might want to try with something simpler
    Dim errcode As Long
    Dim localFileName As String
    
    localFileName = "D:\MyFile.pdf"
    errcode = URLDownloadToFile(0, URL, localFileName, 0, 0)

    If errcode = 0 Then
        MsgBox "Download ok"
    Else
        MsgBox "Error while downloading"
    End If

    ie.Quit
    Set ie = Nothing
End Sub

Private Sub RunBatchFile(strFullBatchFilePath As String)
'Run *.bat files from within VBA
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run Chr(34) & strFullBatchFilePath & Chr(34), 0
    Set WshShell = Nothing
End Sub
