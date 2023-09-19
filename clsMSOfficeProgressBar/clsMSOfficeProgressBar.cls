VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMSOfficeProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File:      clsMSOfficeProgressBar
'Author:    Greg Harward
'Contact:   gharward@gmail.com
'Date:      1/23/16
'
'Summary:
'Class designed to be used to implement the MS Office Applications Excel, Project, Visio, Access - 2003  & 2007 status bar.
'Can be used for other Office applications with minor modification.
'Sample code at end to work with themes.
'PBM_SETMARQUEE style progress not supported as it requires comctl32.dll v6 that is loaded via preinstalled manifest file.
'
'Online References:
'http://www.sweetpotatosoftware.com/SPSBlog/PermaLink,guid,9e494312-e24a-4063-ae8c-ec194db8549a.aspx
'http://hyper.sunjapan.com.cn/~hz/win32/commctl2.htm
'http://msdn2.microsoft.com/en-us/library/Bb264566.aspx
'http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69132&lngWId=1 'Solid VB Source to owner draw themes on status bar.
'
'Revisions:
'Date     Initials    Description of changes

''Implementation Sample:
'Public Sub ProgressExample()
'    Dim obj As clsMSOfficeProgressBar
'    Set obj = New clsMSOfficeProgressBar
'
'    Call obj.UpdateProgress(0, "Processing", , True)
'    Application.Wait (Now + TimeValue("0:00:01"))
'
'    Call obj.UpdateProgress(30, "Processing", , True)
'    Application.Wait (Now + TimeValue("0:00:01"))
'
'    Call obj.UpdateProgress(60, "Processing", , True)
'    Application.Wait (Now + TimeValue("0:00:01"))
'
'    Call obj.UpdateProgress(90, "Processing", , True)
'    Application.Wait (Now + TimeValue("0:00:01"))
'
'    Set obj = Nothing
'End Sub

'//////////////////

''WIN32 API CODE

''Function declares
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetNextWindow Lib "user32" Alias "GetWindow" (ByVal hWnd As Long, ByVal wFlag As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateStatusWindow Lib "comctl32.dll" (ByVal style As Long, ByVal lpszText As String, ByVal hWndParent As Long, ByVal wID As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function GetCurrentThemeName Lib "uxtheme.dll" (ByVal pszThemeFileName As String, ByVal dwMaxNameChars As Integer, ByVal pszColorBuff As String, ByVal cchMaxColorChars As Integer, ByVal pszSizeBuff As String, ByVal cchMaxSizeChars As Integer) As Long

Private Const VER_PLATFORM_WIN32_NT = 2
'Private Const VER_PLATFORM_WIN32_WINDOWS = 1
'Private Const VER_PLATFORM_WIN32s = 0


Private Const WS_VISIBLE = &H10000000
Private Const WS_CHILD = &H40000000
'Private Const WS_POPUP = &H80000000
Private Const WS_CLIPSIBLINGS = &H4000000
'Private Const WS_BORDER = &H800000
Private Const SBARS_SIZEGRIP = &H100

Private Const SW_HIDE = 0
Private Const SW_NORMAL = 1
Private Const SW_MINIMIZE = 6

''Progress Bar Consts
Private Const STATUSCLASSNAMEA = "msctls_statusbar32"   ' status bar control name define
Private Const PROGRESS_CLASSA = "msctls_progress32"   ' progress bar control name define
Private Const WM_USER = &H400                        ' 0x0400 'used by applications to define private messages

'Style
'http://msdn2.microsoft.com/en-us/library/bb760820.aspx
Private Const PBS_SMOOTH = &H1                       ' 0x01 Bar is smooth not segmented
Private Const PBS_VERTICAL = &H4                     ' 0x04 Bar runs up & down
Private Const PBS_MARQUEE = &H8                     ' Marquee mode rather than progress.  Only works with in XP or later. Comctl32.dll version 6
Private Const PBS_SMOOTHREVERSE = &H10                 ' Only works with in XP or later.  Only works when progress bar is horizontal.

'State
Private Const PBST_NORMAL = &H1                     'Green      'Only works with in XP or later. Comctl32.dll version 6
Private Const PBST_ERROR = &H2                          'Red        'Only works with in XP or later. Comctl32.dll version 6
Private Const PBST_PAUSED = &H3                         'Yellow     'Only works with in XP or later. Comctl32.dll version 6
        
Private Const CCM_FIRST = &H2000&
Private Const CCM_SETBKCOLOR = CCM_FIRST + 1

'Progress Bar Messages
Private Const PBM_SETRANGE = WM_USER + 1             ' Set min/max values of progress bar.
Private Const PBM_SETPOS = WM_USER + 2               ' Set current position of the progress bar.
'Private Const PBM_DELTAPOS = WM_USER + 3             ' Advances the position of a progress bar by the specified increment and redraws the control so that the user sees the new position.
'Private Const PBM_SETSTEP = WM_USER + 4              ' Specify step increment of progress bar for PBM_STEP_IT.  Default is 10.
'Private Const PBM_STEPIT = WM_USER + 5               ' Advances the current position of the progress bar.
'Private Const PBM_SETRANGE32 = WM_USER + 6           ' Sets the range of a progress bar control to a 32-bit value.
'Private Const PBM_GETRANGE = WM_USER + 7
'Private Const PBM_GETPOS = WM_USER + 8
Private Const PBM_SETBARCOLOR = WM_USER + 9
'Private Const PBM_SETMARQUEE = WM_USER + 10          ' Marquee mode rather than progress.  Only works with in XP or later.
Private Const PBM_SETSTATE = WM_USER + 16
Private Const PBM_SETBKCOLOR = CCM_SETBKCOLOR
'Private Const PBM_SETTEXTCOLOR = WM_USER + 102     'May have to owner draw to accomplish
'Private Const PBM_SETTEXTBKCOLOR = WM_USER + 104   'May have to owner draw to accomplish

'Status Bar Messages
Private Const SB_SETTEXT As Long = WM_USER + 1          'Set text of status bar window.
'Private Const SB_SIMPLEID = &HFF                        'Specifies that the statusbar is of type simple, so only has one section. '255
'Private Const SB_GETTEXT As Long = WM_USER + 2
'Private Const SB_GETTEXTLENGTH As Long = WM_USER + 3
Private Const SB_SETPARTS As Long = WM_USER + 4
'Private Const SB_GETPARTS As Long = WM_USER + 6
'Private Const SB_GETBORDERS As Long = WM_USER + 7
'Private Const SB_SETMINHEIGHT As Long = WM_USER + 8
'Private Const SB_SIMPLE As Long = WM_USER + 9
'Private Const SB_GETRECT As Long = WM_USER + 10
'Private Const SB_ISSIMPLE As Long = WM_USER + 14
Private Const SB_SETBKCOLOR = CCM_SETBKCOLOR

'Private Const SBB_HORIZONTAL As Long = 0 'horz border width
'Private Const SBB_VERTICAL As Long = 1   'vert border width
'Private Const SBB_DIVIDER As Long = 2    'vert part divider width

'SetWindowPos: Z-Order and windows position messages
'http://msdn2.microsoft.com/en-us/library/ms633545.aspx
'Private Const HWND_BOTTOM = 1
'Private Const HWND_NOTOPMOST = -2
Private Const HWND_TOP = 0                   'Top
'Private Const HWND_TOPMOST = -1  'Top even if deactivated.

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'Private Const SWP_NOACTIVATE = &H10
'Private Const SWP_SHOWWINDOW = &H40
'Private Const SWP_FRAMECHANGED = &H20
'Private Const SWP_NOOWNERZORDER = &H200
'Private Const SWP_NOREDRAW = &H8
'Private Const SWP_NOZORDER = &H4
'Private Const SWP_HIDEWINDOW = &H80

Private Const SBT_SUNKEN = &H0
'Private Const SBT_OWNERDRAW = &H1000

'MS Excel Constants
Private Const xlWait = 2
Private Const xlDefault = &HFFFFEFD1

''Handle for MS Project Application UI Class Objects
'Private Const prjObjFormulaBar = "JWinproj-WhimperEntryBar"                 'Formulabar
Private Const prjObjMain = "JWinproj-WhimperMainClass"                       'Project Application

Private Const ICC_USEREX_CLASSES = &H200

Private Const GWL_HINSTANCE = (-6)
'Private Const GWL_ID = (-12)
'Private Const GWL_STYLE = (-16)

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

'Private Const CCS_TOP As Long = &H1&
'Private Const CCS_NOMOVEY As Long = &H2&
Private Const CCS_BOTTOM As Long = &H3&

'Private Const COLOR_ACTIVECAPTION = 2
'Private Const COLOR_INACTIVECAPTION = 3
Private Const COLOR_MENU = 4
Private Const COLOR_WINDOW = 5
'Private Const COLOR_GRADIENTACTIVECAPTION = 27
'Private Const COLOR_GRADIENTINACTIVECAPTION = 28

Private Enum eMSOfficeApplication
    MSExcel
    MSProject
    MSVisio
    MSAccess
    MSPowerPoint
End Enum

'Registry enumerations
Private Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type tagInitCommonControlsEx
    lSize As Long
    lICC As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type OSVERSIONINFO
    OSVSize         As Long         'size, in bytes, of this data structure
    dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
    dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
    dwBuildNumber   As Long         'NT: build number of the OS 'Win9x: build number of the OS in low-order word. High-order word contains major & minor ver nos.
    PlatformID      As Long         'Identifies the operating system platform.
    szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3" 'Win9x: string providing arbitrary additional information
End Type

'Progress bar class variables
Private hWndProgress As Long
Private hWndStatusBarNew As Long
Private dwStyle As Long
Private xOffset As Long
Private yOffset As Long
Private nWidth As Long
Private nHeight As Long
Private lResult As Long
Private lBackGroundColor As Long
Private lProgressBackgroundColor As Long
'Private lProgressBarColor As Long
Private hWndApplication As Long
Private hInstance As Long
Private strDefaultStatusBarText As String
Private bOriginalStatusBarState As Boolean
Private strOriginalStatusBarText As String
Private MSOfficeProduct As eMSOfficeApplication
Private MSOfficeVersion As Long
Private bValidMSOfficeApplication As Boolean
Private m_Progress As Long
Private m_CurrentProgress As Long
Private m_CurrentProgressText As String

Private vApplication As Object     'Used so that code will compile in any MSO application.

Private Sub Class_Initialize()
'Initialize status bar for MSOffice applications.
    Dim strClassName As String
    Dim yPartWidths(1) As Long
    Dim rWindow As RECT
    
    Set vApplication = Application

    If InitCommonControlsVB = True And Application.Visible = True Then
        
        Select Case Application.name 'Application.Value
            Case "Microsoft Excel"
                MSOfficeProduct = eMSOfficeApplication.MSExcel
            Case "Microsoft Project"
                MSOfficeProduct = eMSOfficeApplication.MSProject
            Case "Microsoft Visio"
                MSOfficeProduct = eMSOfficeApplication.MSVisio
            Case "Microsoft Access"
                MSOfficeProduct = eMSOfficeApplication.MSAccess
            Case "Microsoft PowerPoint"
                MSOfficeProduct = eMSOfficeApplication.MSPowerPoint
            Case Else
                Debug.Print "Unsupported MS Office Product."
                Exit Sub
        End Select
        
        bValidMSOfficeApplication = True
        MSOfficeVersion = val(Application.Version) '2003 = v11, 2007 = v12
        
        xOffset = 4 '8 'Space between right side of progress bar and right side of status bar
        yOffset = 4 '5 'Space from top of status bar to top of progress bar
        nWidth = 240 '232 'Width of progress bar
        nHeight = 16 '12 'Height of progress bar
                
        'Set up application
        Select Case MSOfficeProduct
            Case eMSOfficeApplication.MSExcel
                ' Set the Status Bar to indicate actions to user
                bOriginalStatusBarState = vApplication.DisplayStatusBar 'This could get reset if called again after statusbar is switched on.
                strOriginalStatusBarText = vApplication.StatusBar
                vApplication.DisplayStatusBar = True
                vApplication.Cursor = xlWait         'Enable in Excel
                vApplication.EnableEvents = False    'Stop events from firing to avoid event loops.
                vApplication.ScreenUpdating = False
            
                lBackGroundColor = GetSysColor(COLOR_WINDOW)
                lProgressBackgroundColor = GetSysColor(COLOR_WINDOW) 'GetBackGroundColor
                
'                lBackGroundColor = GetBackGroundColor
'                lProgressBackgroundColor = GetSysColor(COLOR_WINDOW)
                hWndApplication = vApplication.hWnd
                hInstance = vApplication.hInstance
            
            Case MSProject  'Implements new status bar with progress bar on top.
                ' Set the Status Bar to indicate actions to user
                bOriginalStatusBarState = vApplication.DisplayStatusBar 'This could get reset if called again after statusbar is switched on.
                strOriginalStatusBarText = vApplication.StatusBar
                vApplication.DisplayStatusBar = True
                
                lBackGroundColor = GetBackGroundColor
                lProgressBackgroundColor = GetSysColor(COLOR_WINDOW)
                hWndApplication = FindWindowEx(vbEmpty, vbEmpty, prjObjMain, vApplication.Caption) 'MS Project Main 2003 'Caption used to ensure correct window.
                hInstance = GetWindowLong(hWndApplication, GWL_HINSTANCE)
                
            Case MSVisio  'Implements new status bar with progress bar on top.
                ' Set the Status Bar to indicate actions to user
                bOriginalStatusBarState = vApplication.ShowStatusBar 'This could get reset if called again after statusbar is switched on.
                vApplication.ShowStatusBar = True
                
                lBackGroundColor = GetBackGroundColor
                lProgressBackgroundColor = GetSysColor(COLOR_WINDOW)
                hWndApplication = vApplication.WindowHandle32
                hInstance = vApplication.InstanceHandle32    'GetWindowLong(hwndApplication, GWL_HINSTANCE)
            
            Case MSAccess
                lBackGroundColor = GetBackGroundColor
                lProgressBackgroundColor = GetSysColor(COLOR_WINDOW)
                hWndApplication = vApplication.hWndAccessApp
                hInstance = GetWindowLong(hWndApplication, GWL_HINSTANCE)
            
            Case MSPowerPoint
                lBackGroundColor = GetBackGroundColor
                lProgressBackgroundColor = GetSysColor(COLOR_WINDOW)
'                hWndApplication = vApplication.hWndAccessApp
                
                Select Case MSOfficeVersion
                    Case 8
                        strClassName = "PP97FrameClass"
                        hWndApplication = FindWindowEx(vbEmpty, vbEmpty, strClassName, vApplication.Caption & " - [" & vApplication.ActivePresentation.name & "]")  'Caption used to ensure correct window.
                    Case 10, 11
                        strClassName = "PP" & MSOfficeVersion & "FrameClass"
                        hWndApplication = FindWindowEx(vbEmpty, vbEmpty, strClassName, vApplication.Caption)  'Caption used to ensure correct window.
                    Case 12
                        strClassName = "PP" & MSOfficeVersion & "FrameClass"
                        hWndApplication = FindWindowEx(vbEmpty, vbEmpty, strClassName, vApplication.Caption & " - [" & vApplication.ActivePresentation.name & "]")  'Caption used to ensure correct window.
                    Case 14 '2010
                        strClassName = "PPTFrameClass"
                        hWndApplication = FindWindowEx(vbEmpty, vbEmpty, strClassName, vApplication.Caption)  'Caption used to ensure correct window.
                End Select
            
                hInstance = GetWindowLong(hWndApplication, GWL_HINSTANCE)
        End Select
        
        'Delete any previously orphaned status windows
        Call DeleteOrphanedChildWindowsByClass(hWndApplication, STATUSCLASSNAMEA)
        
        'Specifying CCS_BOTTOM here causes SBARS_SIZEGRIP to be invisible by default.
        dwStyle = WS_CHILD Or WS_VISIBLE Or CCS_BOTTOM 'Or CCS_TOP Or SBARS_SIZEGRIP 'Or CCS_NOMOVEY

        hWndStatusBarNew = CreateWindowEx(vbEmpty, STATUSCLASSNAMEA, strDefaultStatusBarText, dwStyle, vbEmpty, vbEmpty, vbEmpty, vbEmpty, hWndApplication, vbEmpty, hInstance, ByVal vbEmpty)
        
        Call GetClientRect(hWndStatusBarNew, rWindow)
        
        'Added Status bar sections to remove repaint flicker.
        yPartWidths(0) = rWindow.Right - nWidth - (2 * xOffset)
        yPartWidths(1) = -1 '-1-the rest
        lResult = SendMessage(hWndStatusBarNew, SB_SETPARTS, 2, yPartWidths(0))
        lResult = SendMessage(hWndStatusBarNew, SB_SETBKCOLOR, vbEmpty, ByVal lBackGroundColor)
        lResult = SetWindowPos(hWndStatusBarNew, HWND_TOP, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE) 'Keep on top to avoid flicker or repainting.
        
        dwStyle = WS_CHILD Or WS_VISIBLE Or WS_CLIPSIBLINGS
        If MSOfficeVersion >= 12 Then '- Makes progress smooth looking.  Office 2007 has smooth looking theme. Excel 2007 = 12.
            dwStyle = dwStyle Or PBS_SMOOTH Or PBS_SMOOTHREVERSE 'Or PBS_MARQUEE
        End If
        
        hWndProgress = CreateWindowEx(vbEmpty, PROGRESS_CLASSA, vbNullString, dwStyle, yPartWidths(0) + xOffset, yOffset, nWidth + 2, nHeight, hWndStatusBarNew, vbEmpty, hInstance, ByVal vbEmpty)
        
        If MSOfficeVersion >= 12 Then '-Set to default color. Office 2007 has smooth looking theme. Excel 2007 = 12.
            lResult = SendMessage(hWndProgress, PBM_SETSTATE, PBST_NORMAL, 0)
        End If
        
        lResult = ShowWindow(hWndProgress, SW_HIDE) 'Initially hide
        lResult = SendMessage(hWndProgress, PBM_SETRANGE, vbEmpty, ByVal MAKELPARAM(0, 100))
        lResult = SendMessage(hWndProgress, PBM_SETBARCOLOR, vbEmpty, ByVal GetBackGroundColor) 'Inherits from theme, so don't have to set.
        lResult = SendMessage(hWndProgress, PBM_SETBKCOLOR, vbEmpty, ByVal lProgressBackgroundColor) 'CLR_DEFAULT)
    ''    lResult = SendMessage(hWndProgress, PBM_SETSTEP, CInt(20), vbEmpty)                 'Set LED Width for PBM_STEPIT
    ''    lResult = ShowWindow(hWndProgress, SW_NORMAL)
    End If
End Sub

Public Sub UpdateProgress(Optional ByVal strStatusBarText As String = "Ready", Optional ByVal lProgress As Single = 0, Optional bAsDecimal As Boolean = False, Optional bAppendProgressToText As Boolean = False)
    If bValidMSOfficeApplication = True And InitCommonControlsVB = True And Application.Visible = True Then
        
        'Move statusbar to fit (possible new) text
        If bAsDecimal = True Then
            lProgress = lProgress * 100
        End If
        
        'Internally progress range is set to 0-100
        If lProgress > 100 Then
            lProgress = 100
        ElseIf lProgress < 0 Then
            lProgress = 0
        End If

        If bAppendProgressToText = True Then
            strStatusBarText = strStatusBarText & VBA.Space(1) & lProgress & "%"
        End If
        
        lResult = SendMessage(hWndProgress, PBM_SETPOS, ByVal lProgress, vbEmpty) 'Update Progress
        m_CurrentProgress = lProgress
        
        If m_Progress <> lProgress Then
            lResult = ShowWindow(hWndProgress, SW_NORMAL) 'Set to visible on first update
            m_Progress = lProgress
        End If
        
'        If StrComp(m_CurrentProgressText, strStatusBarText) <> False Then 'Only change if text changes, reduces flicker visible in progress bar.
            lResult = SendMessage(hWndStatusBarNew, SB_SETTEXT, ByVal SBT_SUNKEN, ByVal CStr(VBA.Space(2) & strStatusBarText))
            m_CurrentProgressText = strStatusBarText
'        End If
        DoEvents
    End If
End Sub

Public Sub IncrementProgress(ByVal strStatusBarText As String, Optional lMaxProgress As Long = 100, Optional ByVal Increment As Long = 1, Optional bAppendProgressToText As Boolean = False)
    If bValidMSOfficeApplication = True And InitCommonControlsVB = True And Application.Visible = True Then
        Dim lProgress As Single
'        'Move statusbar to fit (possible new) text
'        If bAsDecimal = True Then
'            lProgress = lProgress * 100
'        End If
        
        'Internally progress range is set to 0-100
'        If lMaxProgress > 100 Then
'            lMaxProgress = 100
'        ElseIf lMaxProgress < 1 Then
'            lMaxProgress = 1
'        End If
        
        'Get progress as percentage of 100 which is internal max of progress bar.
        lProgress = Round(((m_CurrentProgress + Increment) / lMaxProgress) * 100)
        
        If lProgress > 100 Then
'            lProgress = m_CurrentProgressRelative
            m_CurrentProgress = lMaxProgress
            lProgress = 100
        Else
            m_CurrentProgress = m_CurrentProgress + Increment
        End If
        
        If bAppendProgressToText = True Then
            strStatusBarText = strStatusBarText & VBA.Space(1) & lProgress & "%"
        End If
        
        strStatusBarText = VBA.Space(2) & strStatusBarText

        If m_Progress <> lProgress Then
            lResult = SendMessage(hWndProgress, PBM_SETPOS, ByVal lProgress, vbEmpty) 'Update Progress
            m_Progress = lProgress
        End If

        lResult = ShowWindow(hWndProgress, SW_NORMAL) 'Set to visible on first update
        
'        If StrComp(m_CurrentProgressText, strStatusBarText) <> False Then 'Only change if text changes, reduces flicker visible in progress bar.
            lResult = SendMessage(hWndStatusBarNew, SB_SETTEXT, ByVal SBT_SUNKEN, ByVal strStatusBarText)
            m_CurrentProgressText = strStatusBarText
'        End If
        DoEvents
    End If
End Sub

Private Sub Class_Terminate()
    If bValidMSOfficeApplication = True And vApplication.Visible = True Then
        '// Clean-up.....
        If hWndProgress <> 0 Then
            Call DestroyWindow(hWndProgress)
            hWndProgress = vbEmpty
        End If
        
        If hWndStatusBarNew <> 0 Then
            Call DestroyWindow(hWndStatusBarNew)
            hWndStatusBarNew = vbEmpty
        End If
        
        Select Case MSOfficeProduct
            Case eMSOfficeApplication.MSExcel
                ' Reset the Status bar to its original condition
                'Application.ScreenUpdating = True 'Possibly add
                vApplication.DisplayStatusBar = bOriginalStatusBarState
                vApplication.StatusBar = False  '"Ready" 'strOriginalStatusBarText
                vApplication.Cursor = xlDefault    'Enable in Excel
                vApplication.Visible = True
                vApplication.ScreenUpdating = True
                vApplication.EnableEvents = True
                vApplication.DisplayAlerts = True 'Reenable alerts in the event that they were left off by accident.
                'Err.Clear          'May want to leave this so that setting up an error message after this call is possible using this number.
            Case eMSOfficeApplication.MSProject
                ' Reset the Status bar to its original condition
                'Application.ScreenUpdating = True 'Possibly add
                vApplication.DisplayStatusBar = bOriginalStatusBarState
                vApplication.StatusBar = False  '"Ready"   'strOriginalStatusBarText
                vApplication.Visible = True
                vApplication.ScreenUpdating = True
                vApplication.DisplayAlerts = True 'Reenable alerts in the event that they were left off by accident.
                'Err.Clear          'May want to leave this so that setting up an error message after this call is possible using this number.
            Case eMSOfficeApplication.MSVisio
                ' Reset the Status bar to its original condition
                'Application.ScreenUpdating = True 'Possibly add
                vApplication.ShowStatusBar = bOriginalStatusBarState
'                Application.StatusBar = strOriginalStatusBarText
                vApplication.Visible = True
                vApplication.ScreenUpdating = True
                'Err.Clear          'May want to leave this so that setting up an error message after this call is possible using this number.
            Case eMSOfficeApplication.MSAccess
'                Application.Visible = True
        End Select
    End If
End Sub

Private Sub DeleteOrphanedChildWindowsByClass(ByVal hWndParent As Long, strClassName)
'Delete previously orphaned status bar windows by class type.
    Dim test_hwnd As Long
    Dim test_next As Long
    'Find the first window
    Dim lProgress As Long
    Dim lRet As Long
    Dim sClassNameTest As String
    
    test_hwnd = GetWindow(hWndApplication, GW_CHILD)
    
    Do While test_hwnd <> 0
        'Check if the window isn't a child
'        If GetParent(test_hwnd) = hWndApplication Then
            sClassNameTest = VBA.Space(Len(strClassName))
            Call GetClassName(test_hwnd, sClassNameTest, Len(strClassName) + 1)
            
            test_next = GetNextWindow(test_hwnd, GW_HWNDNEXT)
            
            If sClassNameTest = strClassName Then
                Call DestroyWindow(test_hwnd)
'                VBA.Interaction.DoEvents
            End If
            test_hwnd = test_next
        'retrieve the next window
'        test_hwnd = GetWindow(test_hwnd, GW_HWNDNEXT)
'        test_hwnd = GetNextWindow(test_hwnd, GW_HWNDNEXT)
    Loop
End Sub

Private Function InitCommonControlsVB() As Boolean
'Ensure you've linked to ComCtl32.dll
    On Error Resume Next
    Dim iccex As tagInitCommonControlsEx
    ' Ensure CC available:
    With iccex
        .lSize = LenB(iccex)
        .lICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    InitCommonControlsVB = (Err.Number = 0)

    On Error GoTo 0
End Function

Private Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
    'Combines two integers into a long
    MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

Private Function MAKELONG(wLow As Long, wHigh As Long) As Long
    MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Private Function LoWord(dwValue As Long) As Integer
    CopyMemory LoWord, dwValue, 2
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Theme related
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetBackGroundColor() As Long
    If MSOfficeVersion >= 12 Then 'office 2007 or greater, assume forward compatible.
        GetBackGroundColor = GetOfficeColorScheme()
    ElseIf OSSupportsThemes Then 'Win XP or greater. 'Pre 2007 office application status bar area
        GetBackGroundColor = GetThemeColorSimple 'Simplified approach.  Full themes not yet support.
    Else 'default - OS Display Settings
        GetBackGroundColor = GetSysColor(COLOR_MENU)
    End If
End Function

'Theme color code.
Private Function GetOfficeColorScheme() As Long
    Dim lOfficeTheme As Integer
    If MSOfficeVersion >= 14 Then
        lOfficeTheme = GetRegistryDWORDValueWMI(HKEY_CURRENT_USER, "Software\Microsoft\Office\15.0\Common", "Theme")
    Else
        lOfficeTheme = GetRegistryDWORDValueWMI(HKEY_CURRENT_USER, "Software\Microsoft\Office\12.0\Common", "Theme")
    End If
    Select Case lOfficeTheme
        Case 1 ' ThemeBlue
            GetOfficeColorScheme = RGB(197, 220, 248) 'RGB(227, 239, 255)
        Case 2 ' ThemeSilver
            GetOfficeColorScheme = RGB(231, 232, 235) 'RGB(240, 241, 242)
        Case 3 ' ThemeBlack 'Actually need black background with white text, however white text seems to require owner drawing, so stay grey.
            GetOfficeColorScheme = RGB(216, 215, 216) 'RGB(47, 47, 47)
        Case Else 'Registry entry doesn't exist until theme is changed. Set default color.
            If MSOfficeVersion >= 15 Then 'Office 2012
                Select Case MSOfficeProduct
                    Case eMSOfficeApplication.MSExcel
                        GetOfficeColorScheme = RGB(33, 115, 70) 'Green
                    Case eMSOfficeApplication.MSProject
                        GetOfficeColorScheme = RGB(49, 117, 47) 'Dark Green
                    Case eMSOfficeApplication.MSVisio
                        GetOfficeColorScheme = RGB(57, 85, 163) 'Dark Blue
                    Case eMSOfficeApplication.MSAccess
                        GetOfficeColorScheme = RGB(164, 55, 58) 'Maroon
                    Case eMSOfficeApplication.MSPowerPoint
                        GetOfficeColorScheme = RGB(210, 71, 38) 'Orange
                    Case Else
                        GetOfficeColorScheme = GetSysColor(COLOR_MENU)
                End Select
            ElseIf MSOfficeVersion = 14 Then 'Office 2010
                GetOfficeColorScheme = RGB(216, 222, 228) 'Silver
            Else 'Office 2007
                GetOfficeColorScheme = RGB(197, 220, 248)
            End If
    End Select
End Function

Private Function GetRegistryDWORDValueWMI(hRegTreeKey As hKey, sSubKeyName As String, sValueName As String) As Long
'http://msdn2.microsoft.com/en-us/library/aa394600.aspx
    On Error Resume Next
    
    Dim oReg As Object 'SWbemObjectEx
    Dim uValue As Long ' Integer
    Dim strComputer As String
    Dim lRet As Long
    
    strComputer = "."
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
    lRet = oReg.GetDWORDValue(hRegTreeKey, sSubKeyName, sValueName, uValue)
    
    If lRet = False Then
        GetRegistryDWORDValueWMI = uValue
    End If
    Err.Clear
End Function

Private Function OSSupportsThemes() As Boolean 'IsWinXP
'Private Function IsWinXP() As Boolean
'Modified from Source: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69132&lngWId=1
' IsWinXP = (OSV.PlatformID = VER_PLATFORM_WIN32_NT) And (OSV.dwVerMajor = 5 And OSV.dwVerMinor = 1) And (OSV.dwBuildNumber >= 2600)
    Dim OSV As OSVERSIONINFO

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) <> False Then
        If OSV.PlatformID >= VER_PLATFORM_WIN32_NT Then
            If val(OSV.dwVerMajor & "." & OSV.dwVerMinor) >= 5.1 Then
                OSSupportsThemes = True
'                If OSV.dwBuildNumber > 2600 Then
'                End If
                Exit Function
            End If
        End If
    End If
'Func_ErrHandler:
'    Err.Raise Err.Number, "clsMSOfficeProgressBar.IsWinXP", Err.Description, Err.HelpFile, Err.HelpContext
End Function

Private Function GetThemeColorSimple() As Long
'Modified from Source: http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=69132&lngWId=1
'Possibly replace this approach with API call: GetThemeColor()
'http://msdn.microsoft.com/en-us/library/bb773397(VS.85).aspx
    Dim sFileName As String: sFileName = VBA.Space(255)
    Dim sTempName As String: sTempName = VBA.Space(255)
    Dim sThemeName As String: sThemeName = "None" 'Windows Classic Style
    
    If OSSupportsThemes Then
        If GetCurrentThemeName(sFileName, 255, sTempName, 255, vbNullString, 0) = &H0 Then
            sTempName = VBA.Trim(Replace(VBA.Trim(sTempName), vbNullChar, Empty))
            If sTempName <> vbNullString Then
                sThemeName = sTempName
            End If
        End If
    End If

    Select Case sThemeName
        Case "HomeStead"  'Olive Green
            GetThemeColorSimple = RGB(239, 236, 223)
        Case "Metallic"    'Silver
            GetThemeColorSimple = RGB(233, 224, 229)
        Case "NormalColor"    'Default (blue)
            GetThemeColorSimple = RGB(236, 233, 216)
        Case Else 'None
            GetThemeColorSimple = GetSysColor(COLOR_MENU)
    End Select
End Function
