Attribute VB_Name = "modAPI"
Option Explicit

'ALL WIN32 API CODE
'Pass in long values to these functions using "&" such as "0&" to any function that requires a long, but for which an actual value is being passed.

'Function declares
'Use instead of SendMessage so that application doesn't hang waiting for response if destination window is unresponsive.
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long 'Left, Top, Right, Bottom
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Public Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long     'In screen coordinates.
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long 'In pixel screen coordinates.
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetTextExtentPoint32 Lib "GDI32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As POINTAPI) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Public Declare Function Putfocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Public Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'SetWaitableTimer might be better
'Public Declare Function SetWaitableTimer Lib "kernel32.dll" (ByVal hTimer As Long, ByRef lpDueTime As LARGE_INTEGER, ByVal lPeriod As Long, ByRef pfnCompletionRoutine As PTIMERAPCROUTINE, lpArgToCompletionRoutine As Any, ByVal fResume As Long) As Long 'http://support.microsoft.com/kb/231298
Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function DrawFrameControl Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal un1 As Long, ByVal un2 As Long) As Long
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Public Declare Function GetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function GetFocus Lib "user32" () As Long
Public Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetComboBoxInfo Lib "user32" (ByVal hwndCombo As Long, CBInfo As COMBOBOXINFO) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CompareFileTime Lib "kernel32" (lpFileTime1 As FILETIME, lpFileTime2 As FILETIME) As Long
Public Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Public Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
Public Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetUserNameEx Lib "secur32.dll" Alias "GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long
Public Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Public Declare Function MonitorFromRect Lib "user32.dll" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Public Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MonitorInfo) As Long
Public Declare Function MoveWindow Lib "user32" (ByVal hwnd As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare Function OpenThemeData Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszClassList As Long) As Long
Public Declare Function GetThemeInt Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, piVal As Long) As Long
Public Declare Function GetThemeColor Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal iPartId As Long, ByVal iStateId As Long, ByVal iPropId As Long, ByRef pColor As OLE_COLOR) As Long
Public Declare Function DrawThemeBackground Lib "uxtheme.dll" (ByVal hTheme As Long, ByVal lHDC As Long, ByVal iPartId As Long, ByVal iStateId As Long, pRect As RECT, pClipRect As RECT) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As tagInitCommonControlsEx) As Boolean
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'To get Temp path from environment variable: strResult = Environ("temp") | strResult = Environ("tmp")
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'DC
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SaveDC Lib "GDI32" (ByVal hdc As Long) As Long
Public Declare Function RestoreDC Lib "GDI32" (ByVal hdc As Long, ByVal nSavedDC As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long

'For TwipsPerPixel
'Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
'Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetDeviceCaps Lib "GDI32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'constants
Public Const COLOR_BTNFACE = 15
Public Const COLOR_BTNSHADOW = 16

Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Public Const SM_CXSCREEN = 0 'X Size of screen
Public Const SM_CYSCREEN = 1 'Y Size of Screen
Public Const SM_CXVSCROLL = 2 'X Size of arrow in vertical scroll bar.
Public Const SM_CYHSCROLL = 3 'Y Size of arrow in horizontal scroll bar
Public Const SM_CYCAPTION = 4 'Height of windows caption
Public Const SM_CXBORDER = 5 'Width of no-sizable borders
Public Const SM_CYBORDER = 6 'Height of non-sizable borders
Public Const SM_CXDLGFRAME = 7 'Width of dialog box borders
Public Const SM_CYDLGFRAME = 8 'Height of dialog box borders
Public Const SM_CYVTHUMB = 9 'Height of scroll box on horizontal scroll bar
Public Const SM_CXHTHUMB = 10 ' Width of scroll box on horizontal scroll bar
Public Const SM_CXICON = 11 'Width of standard icon
Public Const SM_CYICON = 12 'Height of standard icon
Public Const SM_CXCURSOR = 13 'Width of standard cursor
Public Const SM_CYCURSOR = 14 'Height of standard cursor
Public Const SM_CYMENU = 15 'Height of menu
Public Const SM_CXFULLSCREEN = 16 'Width of client area of maximized window
Public Const SM_CYFULLSCREEN = 17 'Height of client area of maximized window
Public Const SM_CYKANJIWINDOW = 18 'Height of Kanji window
Public Const SM_MOUSEPRESENT = 19 'True is a mouse is present
Public Const SM_CYVSCROLL = 20 'Height of arrow in vertical scroll bar
Public Const SM_CXHSCROLL = 21 'Width of arrow in vertical scroll bar
Public Const SM_DEBUG = 22 'True if deugging version of windows is running
Public Const SM_SWAPBUTTON = 23 'True if left and right buttons are swapped.
Public Const SM_CXMIN = 28 'Minimum width of window
Public Const SM_CYMIN = 29 'Minimum height of window
Public Const SM_CXSIZE = 30 'Width of title bar bitmaps
Public Const SM_CYSIZE = 31 'height of title bar bitmaps
Public Const SM_CXMINTRACK = 34 'Minimum tracking width of window
Public Const SM_CYMINTRACK = 35 'Minimum tracking height of window
Public Const SM_CXDOUBLECLK = 36 'double click width
Public Const SM_CYDOUBLECLK = 37 'double click height
Public Const SM_CXICONSPACING = 38 'width between desktop icons
Public Const SM_CYICONSPACING = 39 'height between desktop icons
Public Const SM_MENUDROPALIGNMENT = 40 'Zero if popup menus are aligned to the left of the memu bar item. True if it is aligned to the right.
Public Const SM_PENWINDOWS = 41 'The handle of the pen windows DLL if loaded.
Public Const SM_DBCSENABLED = 42 'True if double byte characteds are enabled
Public Const SM_CMOUSEBUTTONS = 43 'Number of mouse buttons.
Public Const SM_CMETRICS = 44 'Number of system metrics
Public Const SM_CLEANBOOT = 67 'Windows 95 boot mode. 0 = normal, 1 = safe, 2 = safe with network
Public Const SM_CXMAXIMIZED = 61 'default width of win95 maximised window
Public Const SM_CXMAXTRACK = 59 'maximum width when resizing win95 windows
Public Const SM_CXMENUCHECK = 71 'width of menu checkmark bitmap
Public Const SM_CXMENUSIZE = 54 'width of button on menu bar
Public Const SM_CXMINIMIZED = 57 'width of rectangle into which minimised windows must fit.
Public Const SM_CYMAXIMIZED = 62 'default height of win95 maximised window
Public Const SM_CYMAXTRACK = 60 'maximum width when resizing win95 windows
Public Const SM_CYMENUCHECK = 72 'height of menu checkmark bitmap
Public Const SM_CYMENUSIZE = 55 'height of button on menu bar
Public Const SM_CYMINIMIZED = 58 'height of rectangle into which minimised windows must fit.
Public Const SM_CYSMCAPTION = 51 'height of windows 95 small caption
Public Const SM_MIDEASTENABLED = 74 'Hebrw and Arabic enabled for windows 95
Public Const SM_NETWORK = 63 'bit o is set if a network is present. Const SM_SECURE = 44 'True if security is present on windows 95 system
Public Const SM_SLOWMACHINE = 73 'true if machine is too slow to run win95.

Public Const GW_OWNER = 4
Public Const GW_HWNDNEXT = 2

Public Const GWL_WNDPROC = (-4)
Public Const GWL_HWNDPARENT = (-8)
Public Const GWL_STYLE = (-16)
Public Const GWL_EXSTYLE = (-20)

Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_NCACTIVATE = &H86
Public Const WM_KILLFOCUS = &H8
Public Const WM_COMMAND = &H111
Public Const WM_KEYUP = &H101
Public Const WM_KEYDOWN = &H100
Public Const WM_CHAR = &H102

Public Const WA_INACTIVE = 0
Public Const WA_ACTIVE = 1

Public Const WH_CALLWNDPROC = 4
Public Const WH_KEYBOARD = 2

Public Const CBN_DROPDOWN = 7

Public Const VK_SHIFT = &H10
Public Const VK_CONTROL = &H11

Public Const KEYEVENTF_EXTENDEDKEY = &H1
Public Const KEYEVENTF_KEYUP = &H2

'http://msdn2.microsoft.com/en-us/library/ms633545.aspx
Public Const HWND_BOTTOM = 1
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0                   'Top
Public Const HWND_TOPMOST = -1  'Top even if deactivated.

Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOZORDER = &H4
Public Const SWP_HIDEWINDOW = &H80

Public Const WS_VISIBLE = &H10000000
Public Const WS_CHILD = &H40000000
Public Const WS_POPUP = &H80000000
Public Const WS_CLIPSIBLINGS = &H4000000

Public Const WS_EX_TOOLWINDOW = &H80

Public Const DFC_BUTTON = 4
Public Const DFCS_BUTTONPUSH = &H10
Public Const DFCS_PUSHED = &H200

Public Const DT_CENTER = &H1

Public Const SW_HIDE = 0
Public Const SW_NORMAL = 1
Public Const SW_MINIMIZE = 6

Public Const SPI_GETWORKAREA = 48

'file creation constants
Public Const GENERIC_WRITE = &H40000000
Public Const GENERIC_READ = &H80000000
Public Const OPEN_EXISTING = 3
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2

'sound related consts
Public Const SND_ASYNC = &H1         '  play asynchronously
Public Const SND_SYNC = &H0
Public Const SND_NODEFAULT = &H2   ' Do not use default sound.
Public Const SND_MEMORY = &H4      ' lpszSoundName points to a
                                         ' memory file.
Public Const SND_NOWAIT = &H2000      '  don't wait if the driver is busy

'Monitor related consts
Public Const MONITOR_DEFAULTTONEAREST = &H2

'Progress Bar Consts
Public Const PROGRESS_CLASS = "msctls_progress32"   ' progress bar control name define
Public Const WM_USER = &H400                        ' 0x0400 'used by applications to define private messages

Public Const PBS_SMOOTH = &H1                       ' 0x01 Bar is smooth not segmented
Public Const PBS_VERTICAL = &H4                     ' 0x04 Bar runs up & down
Public Const m_LED = &H50000000
Public Const m_SMT = &H50000001

Public Const CCM_FIRST = &H2000&
Public Const CCM_SETBKCOLOR = CCM_FIRST + 1

Public Const PBM_SETRANGE = WM_USER + 1             ' Set min/max values of progress bar.
Public Const PBM_SETPOS = WM_USER + 2               ' Set current position of the progress bar.
Public Const PBM_DELTAPOS = WM_USER + 3             ' Advances the position of a progress bar by the specified increment and redraws the control so that the user sees the new position.
Public Const PBM_SETSTEP = WM_USER + 4              ' Specify step increment of progress bar for PBM_STEP_IT.  Default is 10.
Public Const PBM_STEPIT = WM_USER + 5               ' Advances the current position of the progress bar.
Public Const PBM_SETRANGE32 = WM_USER + 6           ' Sets the range of a progress bar control to a 32-bit value.
Public Const PBM_GETRANGE = WM_USER + 7
Public Const PBM_GETPOS = WM_USER + 8
Public Const PBM_SETBARCOLOR = WM_USER + 9
Public Const PBM_SETMARQUEE = WM_USER + 10          ' Marquee mode rather than progress.  Only works with in XP or later.
Public Const PBM_SETBKCOLOR = CCM_SETBKCOLOR

Public Const CLR_DEFAULT = &HFF000000

'strExcelObjClassName
'Handle for Excel Application UI Class Objects       /Descp/
'Retrieve using hwnd = FindWindowEx(FindWindow("XLMAIN", vbNullString), 0&, hWndXLStatusBar, vbNullString)
Public Const xlObjFormulaBar = "EXCEL<"                 'Formulabar
Public Const xlObjMain = "XLMAIN"                       'Excel Application
Public Const xlObjCombarSpace = "EXCEL2"                'Blank space combars
Public Const xlObjStatusBar = "EXCEL4"                  'StatusBar
Public Const xlObjCommandBars = "MsoCommandbar"         'commandbars
Public Const xlObjDesk = "XLDESK"                       'BlankArea sheetMin
Public Const xlObjSheetArea = "EXCEL 7"                 'Sheet Area
Public Const xlObjNameBox = "Combobox"                  'Name box
Public Const xlObjEditNameBox = "Edit"                  'Edit name box
Public Const xlObjScrollBar = "ScrollBar"               'ScrollBars
Public Const xlObjFormulaBarLeft = "EXCEL;"             'Formulabar left
Public Const xlObjPopupDial = "bosa_sdm_XL9"            'Popup Dialogs
Public Const xlObjFontDropList = "Office dropdown"      'Font dropdown
Public Const xlObjSplitBar = "XLCTL"                    'SplitBar

Public Const ICC_PROGRESS_CLASS = &H20

'For TwipsPerPixel X/Y
Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

'For VBA UserForm class name
Private Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'types
Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type SIZE
    Width As Long 'cx
    Height As Long 'cy
End Type

Public Type POINTAPI
    X As Long 'Single
    Y As Long 'Single
End Type

Public Type COMBOBOXINFO
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

Public Type FILETIME 'used for setting time on a file
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WINDOWPOS
    hwnd As Long
    hWndInsertAfter As Long
    X As Long
    Y As Long
    cx As Long
    cy As Long
    flags As Long
End Type

Public Enum EXTENDED_NAME_FORMAT
    NameUnknown = 0
    NameFullyQualifiedDN = 1
    NameSamCompatible = 2
    NameDisplay = 3
    NameUniqueId = 6
    NameCanonical = 7
    NameUserPrincipal = 8
    NameCanonicalEx = 9
    NameServicePrincipal = 10
End Enum

'Monitor related types:
Public Type MonitorInfo
    cbSize As Long
    rcMonitor As RECT
    rcWork As RECT
    dwFlags As Long
End Type

Public Type tagInitCommonControlsEx
   dwSize As Long
   dwICC As Long
End Type

Type LARGE_INTEGER
    lowpart As Long
    highpart As Long
End Type

Public Function MAKELPARAM(wLow As Long, wHigh As Long) As Long
  'Combines two integers into a long
   MAKELPARAM = MAKELONG(wLow, wHigh)
End Function

Public Function MAKELONG(wLow As Long, wHigh As Long) As Long
   MAKELONG = LoWord(wLow) Or (&H10000 * LoWord(wHigh))
End Function

Public Function LoWord(dwValue As Long) As Integer
   CopyMemory LoWord, dwValue, 2
End Function

Function GetTextExtentPoint(hwnd As Long, strText As String) As POINTAPI
    'Get size of strText
    Dim m_hDeviceContext As Long
    Dim ptText As POINTAPI
    Dim lRes As Long
    
    If hwnd <> 0 Then
        'OpenDC
        'GetWindowDC
        m_hDeviceContext = GetDC(hwnd)              ' Get a device context to draw into
        Call SaveDC(m_hDeviceContext)               ' Save state for later
        
        'Get text position
        lRes = GetTextExtentPoint32(m_hDeviceContext, strText, Len(strText), ptText)
        
        'CloseDC
        Call RestoreDC(m_hDeviceContext, -1)        ' Restore the DC state to where we found it
        Call ReleaseDC(hwnd, m_hDeviceContext)      ' Release this DC handle
    
        GetTextExtentPoint = ptText
    End If
End Function

Private Function IsNewComctl32() As Boolean

  'ensures that the Comctl32.dll library is loaded
   Dim icc As tagInitCommonControlsEx
   
   On Error GoTo Err_InitOldVersion
   
   icc.dwSize = Len(icc)
   icc.dwICC = ICC_PROGRESS_CLASS
   
  'VB will generate error 453 "Specified DLL function not found"
  'here if the new version isn't installed
   IsNewComctl32 = InitCommonControlsEx(icc)
   
   Exit Function
   
Err_InitOldVersion:
'   InitCommonControls
End Function

'--------------------------------------------------
Function TwipsPerPixelX() As Single
'--------------------------------------------------
'Returns the width of a pixel, in twips.
'--------------------------------------------------
  Dim lngDC As Long
  lngDC = GetDC(HWND_DESKTOP)
  TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
  ReleaseDC HWND_DESKTOP, lngDC
End Function

'--------------------------------------------------
Function TwipsPerPixelY() As Single
'--------------------------------------------------
'Returns the height of a pixel, in twips.
'--------------------------------------------------
  Dim lngDC As Long
  lngDC = GetDC(HWND_DESKTOP)
  TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
  ReleaseDC HWND_DESKTOP, lngDC
End Function

Public Function PointsPerPixelX(oObject As Object) As Single
'Get UserForm client coordinates and convert points to pixels
'Same as PointsToScreenPixelsX?
    Dim hwndObj As Long
    Dim rRect As RECT
    
    If TypeOf oObject Is UserForm Then
        hwndObj = FindWindow(C_VBA6_USERFORM_CLASSNAME, oObject.Caption)
    Else
        hwndObj = oObject.hwnd
    End If

    Call GetWindowRect(hwndObj, rRect)
    
    If (rRect.Right - rRect.Left) <> 0 Then
        PointsPerPixelX = oObject.Width / (rRect.Right - rRect.Left)
    Else
        PointsPerPixelX = oObject.Width
    End If
End Function

Public Function PointsPerPixelY(oObject As Object) As Single
'Get UserForm client coordinates and convert points to pixels
'Same as PointsToScreenPixelsY?
    Dim hwndObj As Long
    Dim rRect As RECT
    
    If TypeOf oObject Is UserForm Then
        hwndObj = FindWindow(C_VBA6_USERFORM_CLASSNAME, oObject.Caption)
    Else
        hwndObj = oObject.hwnd
    End If

    Call GetWindowRect(hwndObj, rRect)
    
    If (rRect.Right - rRect.Left) <> 0 Then
        PointsPerPixelY = oObject.Height / (rRect.Bottom - rRect.Top)
    Else
        PointsPerPixelY = oObject.Height
    End If
End Function
