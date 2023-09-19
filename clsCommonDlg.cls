VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCommonDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'File:   clsCommonDlg
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        9/4/12

'Common dialog box control for all Office applications. - Works in Office 10 or above.
'Allows for single file selection only.  Supports Window 2k or above.
'Code simultatiously implements two different method in order to achieve the desired functionality.
'The msoFileDialogOpen and msoFileDialogSaveAs constants are not supported in Microsoft Access, so implemented with Win API for these two.
'Implementation table:
'FileDialogFilePicker = 3 'MSO
'FileDialogFolderPicker = 4 'Win API
'FileDialogOpen = 1 'Win API
'FileDialogSaveAs = 2 'Win API

'Includes code OFNHookProc() for centering window, however call back must be in module, so not implemented.

'Sources:
'http://www.mvps.org/access/api/api0001.htm
'http://support.microsoft.com/default.aspx?scid=kb;en-us;288543
'http://www.instantpages.ltd.uk/Office_FileDialog.htm
'http://vbnet.mvps.org/index.html?code/comdlg/filedlgsoverview.htm
'http://vbnet.mvps.org/index.html?code/comdlg/filesavedlg.htm
'http://www.experts-exchange.com/Microsoft/Development/MS_Access/Q_21584944.html

'Sample Implementation:
'Dim strPath As String
'Dim obj As clsCommonDlg
'Set obj = New clsCommonDlg
'strPath = obj.Show(FileDialogFilePicker, ThisWorkbook.FullName, Excel2003)

'===============================================================================================================
'===============================================================================================================
'===============================================================================================================
Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
     'new Win2000 / WinXP members
    pvReserved As Long
    dwReserved As Long
    FlagsEx As Long
End Type

Private Type BROWSEINFO
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MonitorInfo
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

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

Private Enum PathParseMode
    Path
    FileName
    FileExtension
    FileNameWithoutExtension
End Enum

Public Enum eFilterFileType
    Custom = 0 'Supply optional filter
    ProModelModel   'MOD
    ProModelShift   'SFT
    ProModelGraphicsLibrary 'GLB
    ProModelRDB 'RBD
    Excel2003   'XLS
    Excel2007 'XLSX, XLSM, XLSB
    Access2003   'MDB
    Access2007   'ACCDB
    CommaDelimitedText  'CSV
    ExtensibleMarkupLanguage  'XML
    Folder
End Enum

'MsoFileDialogType
Public Enum eMsoFileDialogType
    FileDialogFilePicker = 3    'Win API
    FileDialogFolderPicker = 4  'MSO
    FileDialogOpen = 1          'Win API
    FileDialogSaveAs = 2        'Win API
End Enum

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MonitorInfo) As Long
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
'Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
'Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
'Private Declare Function GetTopWindow Lib "user32" (ByVal hwnd As Long) As Long
'Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wFlag As Long) As Long

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Const WM_INITDIALOG = &H110
Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const MAX_PATH = 260
'Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
'Private Const OFN_SHOWHELP = &H10

Private Const OFN_ENABLEHOOK = &H20
'Private Const OFN_ENABLETEMPLATE = &H40
'Private Const OFN_ENABLETEMPLATEHANDLE = &H80
'Private Const OFN_NOVALIDATE = &H100
'Private Const OFN_ALLOWMULTISELECT = &H200
'Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
'Private Const OFN_SHAREAWARE = &H4000
'Private Const OFN_NOREADONLYRETURN = &H8000
'Private Const OFN_NOTESTFILECREATE = &H10000
'Private Const OFN_NONETWORKBUTTON = &H20000
'Private Const OFN_NOLONGNAMES = &H40000

' New for Windows 95
Private Const OFN_EXPLORER = &H80000
'Private Const OFN_NODEREFERENCELINKS = &H100000
'Private Const OFN_LONGNAMES = &H200000

Private Const MONITORINFOF_PRIMARY = &H1
Private Const MONITOR_DEFAULTTONEAREST = &H2
Private Const MONITOR_DEFAULTTONULL = &H0
Private Const MONITOR_DEFAULTTOPRIMARY = &H1

Private Const GWL_HINSTANCE = (-6)
Private Const prjObjMain = "JWinproj-WhimperMainClass"                       'Project Application

Private m_Application As Object
Private m_hWndApplication As Long
Private m_hInstance As Long

Private Sub Class_Initialize()
    Dim strPPClassName As String
    Set m_Application = Application
    
    Select Case Application.name 'Application.Value
        Case "Microsoft Excel"
            m_hWndApplication = m_Application.hWnd  'hwndApplication = FindWindowEx(vbEmpty, vbEmpty, xlObjMain, Application.Caption) 'Excel Main 2003 'Caption used to ensure correct window.
            m_hInstance = m_Application.hInstance 'hInstance = GetWindowLong(hwndApplication, GWL_HINSTANCE)
        Case "Microsoft Project"
            m_hWndApplication = FindWindowEx(vbEmpty, vbEmpty, prjObjMain, m_Application.Caption) 'MS Project Main 2003 'Caption used to ensure correct window.
            m_hInstance = GetWindowLong(m_hWndApplication, GWL_HINSTANCE)
        Case "Microsoft Visio"
            m_hWndApplication = m_Application.WindowHandle32
            m_hInstance = m_Application.InstanceHandle32    'GetWindowLong(hwndApplication, GWL_HINSTANCE)
        Case "Microsoft Access"
            m_hWndApplication = m_Application.hWndAccessApp
            m_hInstance = GetWindowLong(m_hWndApplication, GWL_HINSTANCE)
        Case "Microsoft PowerPoint"
            Select Case val(Application.Version)
                Case 10, 11, 12
                    strPPClassName = "PP" & val(Application.Version) & "FrameClass"
                Case 8
                    strPPClassName = "PP97FrameClass"
            End Select
            
            m_hWndApplication = FindWindowEx(vbEmpty, vbEmpty, strPPClassName, m_Application.Caption & " - [" & m_Application.ActivePresentation.name & "]")  'MS Project Main 2003 'Caption used to ensure correct window.
            m_hInstance = GetWindowLong(m_hWndApplication, GWL_HINSTANCE)
        Case Else
            Debug.Print "Unsupported MS Office Product."
    '            Call MsgBox("Unsupported MS Office Version.")
    End Select
End Sub

Public Function GetValidFile(ByRef strFilePath As String, FilterFileType As eFilterFileType, strPromptDialogTitle As String, Optional bPromptForUseFile As Boolean = False, Optional ByVal colFileFilters As Collection) As String
'Returns valid file or vbnullstring
    On Error GoTo errsub
    
    Dim strNewFile As String
    Dim bPreviousSetting As Boolean
    
    If m_Application.name = "Microsoft Excel" Then
        bPreviousSetting = m_Application.ScreenUpdating
        m_Application.ScreenUpdating = True
'        If strFilePath = vbNullString Then
'            strFilePath = m_Application.ThisWorkbook.FullName
'        End If
    End If
    
    'Get export file.
    If strFilePath <> vbNullString And DoesFilePathExist(strFilePath) = True And ParsePath(strFilePath, FileName) <> vbNullString Then 'File found
        strNewFile = strFilePath 'Valid file
        If bPromptForUseFile = True Then
            If MsgBox("The following file was found: " & vbCrLf & vbCrLf & strFilePath & vbCrLf & vbCrLf & "Use this file?", vbQuestion Or vbYesNo, strPromptDialogTitle) = vbNo Then
                strNewFile = Me.Show(FileDialogFilePicker, strFilePath, FilterFileType, colFileFilters)
            End If
        End If
    ElseIf strFilePath = vbNullString Then 'File not found
        If MsgBox("No source file found. " & vbCrLf & vbCrLf & "Browse for the file?", vbQuestion Or vbYesNo, strPromptDialogTitle) = vbYes Then
            strNewFile = Me.Show(FileDialogFilePicker, strFilePath, FilterFileType, colFileFilters)
        End If
    Else
        If MsgBox("The following file was not found: " & vbCrLf & vbCrLf & strFilePath & vbCrLf & vbCrLf & "Browse for the file?", vbQuestion Or vbYesNo, strPromptDialogTitle) = vbYes Then
            strNewFile = Me.Show(FileDialogFilePicker, strFilePath, FilterFileType, colFileFilters)
        End If
    End If
    
    'Validate file - ShowCommonDlg() can return empty string
    If strNewFile <> vbNullString And DoesFilePathExist(strNewFile) = True Then
        'Save value to Custom Properties
'                Call SetWorkbookProperty(ThisWorkbook, strCustomPropName, strNewFile, True)
        GetValidFile = strNewFile
    Else
        Err.Raise 53, , "File not found"
    End If

errsub:
    If m_Application.name = "Microsoft Excel" And m_Application.ScreenUpdating <> bPreviousSetting Then
        m_Application.ScreenUpdating = bPreviousSetting 'Setting causes flicker, so only set if was reset above.
    End If
End Function

Public Function GetValidFolder(ByRef strFolderPath As String, strPromptDialogTitle As String, Optional bPromptForUseFolder As Boolean = False) As String
'Returns valid path or vbnullstring
    On Error GoTo errsub
    
    Dim strNewFolder As String
    Dim bPreviousSetting As Boolean
    
    If m_Application.name = "Microsoft Excel" Then
        bPreviousSetting = m_Application.ScreenUpdating
        m_Application.ScreenUpdating = True
        If strFolderPath = vbNullString Or DoesFilePathExist(strFolderPath, True) = False Then
            strFolderPath = m_Application.ThisWorkbook.Path
        End If
    End If
    
    'Get export file.
    If strFolderPath <> vbNullString And DoesFilePathExist(strFolderPath, True) = True And ParsePath(strFolderPath, Path) <> vbNullString Then 'Folder found
        strNewFolder = strFolderPath 'Valid file
        If bPromptForUseFolder = True Then
            If MsgBox("The following folder was found: " & vbCrLf & vbCrLf & strFolderPath & vbCrLf & vbCrLf & "Use this folder?", vbQuestion Or vbYesNo, strPromptDialogTitle) = vbNo Then
                strNewFolder = Me.Show(FileDialogFolderPicker, strFolderPath, Folder)
            End If
        End If
    ElseIf strFolderPath = vbNullString Then 'File not found
        If MsgBox("No source folder found. " & vbCrLf & vbCrLf & "Browse for the folder?", vbQuestion Or vbYesNo, strPromptDialogTitle) = vbYes Then
            strNewFolder = Me.Show(FileDialogFolderPicker, strFolderPath, Folder)
        End If
    Else
        If MsgBox("The following folder was not found: " & vbCrLf & vbCrLf & strFolderPath & vbCrLf & vbCrLf & "Browse for the folder?", vbQuestion Or vbYesNo, strPromptDialogTitle) = vbYes Then
            strNewFolder = Me.Show(FileDialogFolderPicker, strFolderPath, Folder)
        End If
    End If
    
    'Validate file - ShowCommonDlg() can return empty string
    If strNewFolder <> vbNullString And DoesFilePathExist(strFolderPath, True) = True Then
        'Save value to Custom Properties
'                Call SetWorkbookProperty(ThisWorkbook, strCustomPropName, strNewFolder, True)
        GetValidFolder = strNewFolder
    Else
        Err.Raise 53, , "File not found"
    End If

errsub:
    If m_Application.name = "Microsoft Excel" And m_Application.ScreenUpdating <> bPreviousSetting Then
        m_Application.ScreenUpdating = bPreviousSetting 'Setting causes flicker, so only set if was reset above.
    End If
End Function

Public Function Show(dialogType As eMsoFileDialogType, ByVal strPath As String, FilterFileType As eFilterFileType, Optional ByVal colFileFilters As Collection) As String
    Dim oFileDialog As Office.FileDialog 'Object
    Dim OFN As OPENFILENAME
    Dim fResult As Boolean
    Dim hWndApplication As Long
    Dim hInstance As Long
    Dim bPreviousSetting As Boolean
    Dim strTemp As String
    Dim strFolderPath As String
    Dim col As Variant
    
    If colFileFilters Is Nothing Then
        Set colFileFilters = New Collection
    End If
    
    If strPath <> vbNullString Then
        If DoesFilePathExist(strFolderPath, True) = True And DoesFilePathExist(strFolderPath) = False Then 'Path
            If Right(strPath, 1) = "\" Then
                strFolderPath = strPath
            Else
                strFolderPath = strPath & "\"
            End If
        ElseIf DoesFilePathExist(strFolderPath, True) = DoesFilePathExist(strFolderPath) Then  'File
            strFolderPath = ParsePath(strPath, PathParseMode.Path)
        End If
    End If
    
    With OFN
        .lStructSize = Len(OFN) 'Win2k or later supported 'http://vbnet.mvps.org/index.html?code/comdlg/filesavedlg.htm
        .lpstrDefExt = GetFileFilter(colFileFilters, FilterFileType)
        .nFileExtension = Len(.lpstrDefExt) 'DefaultExt
        
        If dialogType = eMsoFileDialogType.FileDialogFilePicker Or dialogType = eMsoFileDialogType.FileDialogOpen Or dialogType = eMsoFileDialogType.FileDialogSaveAs Then
            For Each col In colFileFilters
                .lpstrFilter = .lpstrFilter & IIf(.lpstrFilter <> vbNullString, ",", vbNullString) & col
            Next
            
            'convert MSO sytle filers to API style
            .lpstrFilter = VBA.Replace(.lpstrFilter, ",", vbNullChar)
        ElseIf dialogType = eMsoFileDialogType.FileDialogFolderPicker Then
            Set oFileDialog = Application.FileDialog(dialogType)
'            If dialogType = eMsoFileDialogType.FileDialogFilePicker Then
'                oFileDialog.Filters.Clear
'                For Each col In colFileFilters
'                    strTemp = Split(col, ",")(0)
'                    strTemp = VBA.Trim(VBA.Left(strTemp, InStr(1, strTemp, "(") - 1))
'                    Call oFileDialog.Filters.Add(strTemp, Split(col, ",")(1))
'                Next
'            End If
        End If
    
        .hwndOwner = m_hWndApplication
        .hInstance = m_hInstance

        ' Allocate string space for the returned strings.
        .nFilterIndex = 1
        .nMaxCustFilter = MAX_PATH
        .lpstrCustomFilter = VBA.String(MAX_PATH, vbNullChar) 'VBA.Space(MAX_PATH)
        
        .nMaxFile = MAX_PATH
        .lpstrFile = VBA.String(MAX_PATH, vbNullChar) 'VBA.Space(MAX_PATH) 'strFileName
        
        .nMaxFileTitle = MAX_PATH
        .lpstrFileTitle = VBA.String(MAX_PATH, vbNullChar) 'VBA.Space(MAX_PATH) 'strFileTitle
        
        .lpstrInitialDir = strFolderPath
        'nFileOffset
        'lCustData
        .lpfnHook = 0 'FARPROC(AddressOf OFNHookProc)
        
        If m_Application.name = "Microsoft Excel" Then
            bPreviousSetting = m_Application.ScreenUpdating
            If bPreviousSetting = False Then
                m_Application.ScreenUpdating = True
            End If
        End If
    
        Select Case dialogType
            'Use API version here
            Case eMsoFileDialogType.FileDialogFilePicker
                .lpstrTitle = "Browse File"
                .flags = OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST + OFN_HIDEREADONLY '+ OFN_EXPLORER + OFN_ENABLEHOOK
                If GetOpenFileName(OFN) <> False Then
                    Show = VBA.Trim(VBA.Replace(.lpstrFile, vbNullChar, vbNullString))
                End If
            
            'Use API version here
            Case eMsoFileDialogType.FileDialogOpen
                .lpstrTitle = "File Open"
                .flags = OFN_PATHMUSTEXIST + OFN_FILEMUSTEXIST + OFN_HIDEREADONLY '+ OFN_EXPLORER + OFN_ENABLEHOOK
                If GetOpenFileName(OFN) <> False Then
                    strPath = VBA.Trim(VBA.Replace(.lpstrFile, vbNullChar, vbNullString))
                    Show = strPath
                    Call LoadFile(strPath)
                End If
                
            'Use MsoFileDialogType.FileDialogFolderPicker
            'See also: http://www.codeproject.com/Articles/14454/Extended-Folder-Browser
            Case eMsoFileDialogType.FileDialogFolderPicker
                oFileDialog.InitialFileName = strFolderPath
'                oFileDialog.InitialView = msoFileDialogViewDetails
'                Set oFileDialog = Application.FileDialog(dialogType)
                If oFileDialog.Show = True Then 'User hit OK (not -1)
                    Show = oFileDialog.SelectedItems(1)
                End If
                
'                Here could also use SHBrowseForFolder() 'http://www.developerfusion.com/code/2127/browse-for-folder-vba/
'                Dim lpBrowseInfo As BROWSEINFO
'                lpBrowseInfo.lpszTitle = "Browse Folder"
'                Call SHBrowseForFolder(lpBrowseInfo)
                            
            'Use API version here
            Case eMsoFileDialogType.FileDialogSaveAs
                .lpstrTitle = "File Save"
                .flags = OFN_PATHMUSTEXIST + OFN_HIDEREADONLY + OFN_OVERWRITEPROMPT '+ OFN_EXPLORER + OFN_ENABLEHOOK
                If GetSaveFileName(OFN) <> False Then
                    strPath = VBA.Trim(VBA.Replace(.lpstrFile, vbNullChar, vbNullString))
                    'If exists kill file
                    If DoesFilePathExist(strFolderPath) = True Then
                        Call Kill(strPath)
                    End If
                    Show = strPath
                End If
        End Select
    End With
    
errsub:
    Set oFileDialog = Nothing
    If m_Application.name = "Microsoft Excel" And m_Application.ScreenUpdating <> bPreviousSetting Then
        m_Application.ScreenUpdating = bPreviousSetting
    End If
End Function

Private Function GetFileFilter(colFileFilters As Collection, FilterFileType As eFilterFileType) As String
    Select Case FilterFileType
        Case ProModelModel
            colFileFilters.Add "ProModel model files (*.mod),*.mod"
            GetFileFilter = "mod"
        Case ProModelShift
            colFileFilters.Add "ProModel shift files (*.sft),*.sft"
            GetFileFilter = "sft"
        Case ProModelGraphicsLibrary
            colFileFilters.Add "ProModel graphics library files (*.glb),*.glb"
            GetFileFilter = "glb"
        Case ProModelRDB
            colFileFilters.Add "ProModel RDB (*.rdb),*.rdb"
            GetFileFilter = "rdb"
        Case Excel2003
            colFileFilters.Add "Excel 2003 files (*.xls),*.xls"
            GetFileFilter = "xls"
        Case Excel2007
            colFileFilters.Add "Excel 2007 files (*.xls;*.xlsx;*.xlsm;*.xlsb),*.xls;*.xlsx;*.xlsm;*.xlsb"
            GetFileFilter = "xlsm"
        Case Access2003
            colFileFilters.Add "Access 2003 files (*.mdb),*.mdb"
            GetFileFilter = "mdb"
        Case Access2007
            colFileFilters.Add "Access 2007 files (*.accdb),*.accdb"
            GetFileFilter = "accdb"
        Case CommaDelimitedText
            colFileFilters.Add "Comma-delimited text files (*.csv),*.csv"
            GetFileFilter = "csv"
        Case ExtensibleMarkupLanguage
            colFileFilters.Add "XML files (*.xml),*.xml"
            GetFileFilter = "xml"
        Case Custom
            colFileFilters.Add "All Files (*.*),*.*"
            GetFileFilter = "*" 'File extention if none specified.
        Case Folder
            'Folder dialog
    End Select
End Function

Public Function GetFilterFileTypeByExtension(strFileName As String) As eFilterFileType
    Dim strExtension As String
    
    strExtension = VBA.UCase(ParsePath(strFileName, PathParseMode.FileExtension))
    Select Case strExtension
        Case "MOD"
            GetFilterFileTypeByExtension = eFilterFileType.ProModelModel
        Case "SFT"
            GetFilterFileTypeByExtension = eFilterFileType.ProModelShift
        Case "GLB"
            GetFilterFileTypeByExtension = eFilterFileType.ProModelGraphicsLibrary
        Case "RDB"
            GetFilterFileTypeByExtension = eFilterFileType.ProModelRDB
        Case "XLS"
            GetFilterFileTypeByExtension = eFilterFileType.Excel2003
        Case "XLSM", "XLSB", "XLSX"
            GetFilterFileTypeByExtension = eFilterFileType.Excel2007
        Case "MDB"
            GetFilterFileTypeByExtension = eFilterFileType.Access2003
        Case "ACCDB"
            GetFilterFileTypeByExtension = eFilterFileType.Access2007
        Case "CSV"
            GetFilterFileTypeByExtension = eFilterFileType.CommaDelimitedText
        Case "XML"
            GetFilterFileTypeByExtension = eFilterFileType.ExtensibleMarkupLanguage
        Case Else
            GetFilterFileTypeByExtension = eFilterFileType.Custom
    End Select
End Function

''============================================================================================================
''==Helper Functions
''============================================================================================================
Private Function DoesFilePathExist(strPath As String, Optional IsDirectory As Boolean = False) As Boolean
'Replacement for Dir() function that can return error #52 "Bad file name or number" when testing for invalid file with a network path.
    On Error GoTo errsub:
    
    Dim fso As Object 'FileSystemObject
     
    Set fso = CreateObject("Scripting.Filesystemobject")
    
    If IsDirectory = False Then
        DoesFilePathExist = fso.FileExists(strPath)
    Else
        DoesFilePathExist = fso.FolderExists(strPath)
    End If
errsub:
    Set fso = Nothing
End Function

Private Function OFNHookProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  'On initialization, set aspects of the API dialog that are not obtainable through manipulating the OPENFILENAME structure members.
  'Modified from: http://vbnet.mvps.org/index.html?code/comdlg/filesavedlg.htm
    Dim hWndParent As Long
    Dim rParent As RECT
    Dim rScreen As RECT
    Dim hMonitor As Long
    Dim newLeft As Long
    Dim newTop As Long
    Dim dlgWidth As Long
    Dim dlgHeight As Long
    Dim scrWidth As Long
    Dim scrHeight As Long
    Dim MI As MonitorInfo
    Dim hwndObj As Long
    
    Select Case uMsg
        Case WM_INITDIALOG
             hWndParent = GetParent(hwndObj)
            
             If hWndParent <> 0 Then
                Call GetWindowRect(hWndParent, rParent)
                
                'get the monitor closest to parent window
                hMonitor = MonitorFromWindow(hWndParent, MONITOR_DEFAULTTONEAREST)
                
                If hMonitor <> 0 Then
                    'get this monitor size
                    MI.cbSize = Len(MI)
                    Call GetMonitorInfo(hMonitor, MI)
            
                    rScreen = MI.rcWork 'virtual-screen pixel coordinates
                
                    'convert the pixels into twips
                    rScreen.Left = rScreen.Left * TwipsPerPixelX
                    rScreen.Top = rScreen.Top * TwipsPerPixelY
                    rScreen.Right = rScreen.Right * TwipsPerPixelX
                    rScreen.Bottom = rScreen.Bottom * TwipsPerPixelY
                    
                    rParent.Left = rParent.Left * TwipsPerPixelX
                    rParent.Top = rParent.Top * TwipsPerPixelY
                    rParent.Right = rParent.Right * TwipsPerPixelX
                    rParent.Bottom = rParent.Bottom * TwipsPerPixelY
                
        '                oObject.Left = (rParent.Left + ((rParent.Right - rParent.Left) / 2) - (oObject.Width / 2))
        '                oObject.Top = (rParent.Top + ((rParent.Bottom - rParent.Top) / 2) - (oObject.Height / 2))
                
                    dlgWidth = rParent.Right - rParent.Left
                    dlgHeight = rParent.Bottom - rParent.Top
            
                    scrWidth = rScreen.Right - rScreen.Left
                    scrHeight = rScreen.Bottom - rScreen.Top
        
                    newLeft = (scrWidth - dlgWidth) \ 2
                    newTop = (scrHeight - dlgHeight) \ 2
                    
                    Call MoveWindow(hWndParent, newLeft, newTop, dlgWidth, dlgHeight, True)
                 
                OFNHookProc = 1
            End If
        End If
   End Select
End Function

'--------------------------------------------------
Private Function TwipsPerPixelX() As Single
'--------------------------------------------------
'Returns the width of a pixel, in twips.
'--------------------------------------------------
    Dim lngDC As Long
    lngDC = GetDC(HWND_DESKTOP)
    TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
    ReleaseDC HWND_DESKTOP, lngDC
End Function

'--------------------------------------------------
Private Function TwipsPerPixelY() As Single
'--------------------------------------------------
'Returns the height of a pixel, in twips.
'--------------------------------------------------
    Dim lngDC As Long
    lngDC = GetDC(HWND_DESKTOP)
    TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
    ReleaseDC HWND_DESKTOP, lngDC
End Function

Private Function FARPROC(ByVal pfn As Long) As Long
  'Dummy procedure that receives and returns the return value of the AddressOf operator.
  'Obtain and set the address of the callback.  This workaround is needed as you can't assign AddressOf directly to a member of a user-  'defined type, but you can assign it to another long and use that (as returned here)
   FARPROC = pfn
End Function

Private Sub LoadFile(strFilePath As String, Optional appStyle As nShowCmd = nShowCmd.SW_SHOW)
'Loads an exe or file with the default provider. Returns App Instance.
'http://msdn.microsoft.com/en-us/library/bb762153(VS.85).aspx
    Call ShellExecute(0&, vbNullString, strFilePath, vbNullString, vbNullString, appStyle)
End Sub

Private Function ParsePath(ByVal strPath As String, iMode As PathParseMode) As String
    'Take the path passed in and return the filename, or the path base upon iMode.
    'Path returns with trailing Application.PathSeparator ("\")
    'If file name is passed in to returnPath, empty string is returned.

    On Error GoTo errsub

    Dim fso As Object
    Set fso = CreateObject("Scripting.Filesystemobject")

    Select Case iMode
        Case PathParseMode.FileExtension
            ParsePath = fso.GetExtensionName(strPath) 'File extension name
        Case PathParseMode.FileName
            ParsePath = fso.GetFileName(strPath) 'File name with extension
        Case PathParseMode.Path
            If InStr(1, strPath, "\") > 0 Then  '"\" exists
                strPath = fso.GetParentFolderName(strPath) 'File path
                ParsePath = fso.BuildPath(strPath, "\") 'Add "\"
            End If
        Case PathParseMode.FileNameWithoutExtension
            ParsePath = fso.GetBaseName(strPath) 'File name without extension
    End Select
errsub:
    Set fso = Nothing
End Function

Public Function GetWorkbookProperty(wkbActive As Object, strPropertyName As String, Optional bIsCustomProperty As Boolean = True) As Variant
'http://www.cpearson.com/excel/docprop.htm 'Modified/Simplified
'Check for empty return with IsEmpty().
    On Error Resume Next

    If bIsCustomProperty = True Then
        GetWorkbookProperty = wkbActive.CustomDocumentProperties(strPropertyName).Value
    Else
        GetWorkbookProperty = wkbActive.BuiltinDocumentProperties(strPropertyName).Value
    End If

End Function

Public Sub SetWorkbookProperty(wkbActive As Object, strPropertyName As String, vPropertyValue As Variant, Optional bIsCustomProperty As Boolean = True)
'Private Sub SetProperty(WorkbookName As String, PropName As String, PValue As Variant, PropCustom As Boolean)
    'http://www.cpearson.com/excel/docprop.htm 'Modified/Simplified
    'wkbDestination             Workbook whose property is to be set.
    'strPropertyName            A string containing the name of the property
    'vPropertyValue             A variant containing the value of the property
    'PropCustom                 A boolean indicating whether the property is a Custom Document Property.
    On Error Resume Next

    Dim DocProps As DocumentProperties
    Dim TheType As Long

    If bIsCustomProperty = True Then
        Set DocProps = wkbActive.CustomDocumentProperties
    Else
        Set DocProps = wkbActive.BuiltinDocumentProperties
    End If

    Select Case varType(vPropertyValue)
        Case vbBoolean
            TheType = msoPropertyTypeBoolean
        Case vbDate
            TheType = msoPropertyTypeDate
        Case vbDouble, vbLong, vbSingle, vbCurrency
            TheType = msoPropertyTypeFloat
        Case vbInteger
            TheType = msoPropertyTypeNumber
        Case vbString
            TheType = msoPropertyTypeString
        Case Else
            TheType = msoPropertyTypeString
    End Select

    'If property doesn't already exist add it
    If GetWorkbookProperty(wkbActive, strPropertyName, bIsCustomProperty) = Empty Then
        Call DocProps.Add(strPropertyName, False, TheType, vPropertyValue)
    End If

    DocProps(strPropertyName).Value = vPropertyValue
End Sub
