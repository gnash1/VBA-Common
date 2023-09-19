VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPositionWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File:   clsPositionWindow
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        7/19/12

'Description: Positions window into selected position by calling ForceWindowIntoWorkArea().
'Sample Implementation:
'    Dim oPW As New clsPositionWindow
'    Call oPW.ForceWindowIntoWorkArea(Me, vbStartUpCenterParent)
'    Set oPW = Nothing

Public Enum StartUpPositionConstants 'VBRUN
    vbStartUpManual = 0 'Leave it alone
    vbStartUpCenterParent = 1
    vbStartUpCenterScreen = 2
    vbStartUpWindowsDefault = 3 '0,0
End Enum

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

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type SIZE
    Width As Long 'cx
    Height As Long 'cy
End Type

Private Const MONITORINFOF_PRIMARY = &H1
Private Const MONITOR_DEFAULTTONEAREST = &H2
Private Const MONITOR_DEFAULTTONULL = &H0
Private Const MONITOR_DEFAULTTOPRIMARY = &H1

Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const SM_CXSCREEN = 0
Private Const SM_CYSCREEN = 1

Private Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame" 'For VBA UserForm class name

Private Declare Function GetMonitorInfo Lib "user32.dll" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MonitorInfo) As Long
Private Declare Function MonitorFromWindow Lib "user32.dll" (ByVal hWnd As Long, ByVal dwFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private m_Application As Object
Private m_ApplicationCursor As Long 'XlMousePointer
Private m_ApplicationEnableEvents As Boolean
Private m_ApplicationScreenUpdating As Boolean

Private Sub UserForm_Initialize()
    Dim oActive As Object
    Set m_Application = Application
                                    
End Sub

Private Sub UserForm_Terminate()
    Select Case m_Application.name 'Application.Value
        Case "Microsoft Excel"
            m_Application.Cursor = m_ApplicationCursor
            m_Application.EnableEvents = m_ApplicationEnableEvents
            If m_Application.ScreenUpdating <> m_ApplicationScreenUpdating Then 'Reduce repaint flicker
                m_Application.ScreenUpdating = m_ApplicationScreenUpdating
            End If
        Case "Microsoft Visio"
            m_Application.EventsEnabled = m_ApplicationEnableEvents
    End Select
End Sub

Public Sub ForceWindowIntoWorkArea(oObject As Object, Optional ePosition As StartUpPositionConstants = vbStartUpCenterScreen)
'Public Sub ForceWindowIntoWorkArea(oObject As Object, Optional ePosition As StartUpPositionConstants = vbStartUpCenterParent)
'Modified to work for VBA UserForm control which has units measured in points
'Also works for controls which have units measured in twips.
    On Error GoTo errsub
    Dim MI As MonitorInfo
    Dim rScreen As RECT
    Dim rParent As RECT
    Dim hMonitor As Long
    Dim hwndObj As Long
    Dim hWndParent As Long

    If TypeOf oObject Is UserForm Then
        hwndObj = FindWindow(C_VBA6_USERFORM_CLASSNAME, oObject.Caption)
    Else 'Form
        hwndObj = oObject.hWnd
    End If

    oObject.StartUpPosition = vbStartUpManual

    hWndParent = GetParent(hwndObj)

    Call GetWindowRect(hWndParent, rParent)

    'get the monitor closest to parent window
    hMonitor = MonitorFromWindow(hWndParent, MONITOR_DEFAULTTONEAREST)

    If hMonitor <> 0 Then
        'get this monitor size
        MI.cbSize = Len(MI)
        Call GetMonitorInfo(hMonitor, MI)

        rScreen = MI.rcWork 'virtual-screen pixel coordinates

        If TypeOf oObject Is UserForm Then
            'convert the pixels into points
            rScreen.Left = rScreen.Left * PointsPerPixelX(oObject)
            rScreen.Top = rScreen.Top * PointsPerPixelY(oObject)
            rScreen.Right = rScreen.Right * PointsPerPixelX(oObject)
            rScreen.Bottom = rScreen.Bottom * PointsPerPixelY(oObject)

            rParent.Left = rParent.Left * PointsPerPixelX(oObject)
            rParent.Top = rParent.Top * PointsPerPixelY(oObject)
            rParent.Right = rParent.Right * PointsPerPixelX(oObject)
            rParent.Bottom = rParent.Bottom * PointsPerPixelY(oObject)
        Else
            'convert the pixels into twips
            rScreen.Left = rScreen.Left * TwipsPerPixelX
            rScreen.Top = rScreen.Top * TwipsPerPixelY
            rScreen.Right = rScreen.Right * TwipsPerPixelX
            rScreen.Bottom = rScreen.Bottom * TwipsPerPixelY

            rParent.Left = rParent.Left * TwipsPerPixelX
            rParent.Top = rParent.Top * TwipsPerPixelY
            rParent.Right = rParent.Right * TwipsPerPixelX
            rParent.Bottom = rParent.Bottom * TwipsPerPixelY
        End If

'    Else
'        'failed..so get the number from vb6
'        rScreen.Left = 0
'        rScreen.Top = 0
'        rParent.Left = 0
'        rParent.Top = 0
'
'        If TypeOf oObject Is UserForm Then
'            rScreen.Right = GetScreenInfo.Width * PointsPerPixelX(oObject)
'            rScreen.Bottom = GetScreenInfo.Height * PointsPerPixelY(oObject)
'
'            rParent.Right = GetScreenInfo.Width * PointsPerPixelX(oObject)
'            rParent.Bottom = GetScreenInfo.Height * PointsPerPixelY(oObject)
'        Else
'            rScreen.Right = GetScreenInfo.Width * TwipsPerPixelX
'            rScreen.Bottom = GetScreenInfo.Height * TwipsPerPixelY
'
'            rParent.Right = GetScreenInfo.Width * TwipsPerPixelX
'            rParent.Bottom = GetScreenInfo.Height * TwipsPerPixelY
'        End If

        Select Case ePosition
            Case StartUpPositionConstants.vbStartUpManual 'Leave it
                'Force the window within the area
                Call AdjustWindowPositionIntoWorkArea(oObject, rScreen)

            Case StartUpPositionConstants.vbStartUpCenterParent
                oObject.Left = (rParent.Left + ((rParent.Right - rParent.Left) / 2) - (oObject.Width / 2))
                oObject.Top = (rParent.Top + ((rParent.Bottom - rParent.Top) / 2) - (oObject.Height / 2))

                'Force the window within the area
                Call AdjustWindowPositionIntoWorkArea(oObject, rScreen)

            Case StartUpPositionConstants.vbStartUpCenterScreen
                oObject.Left = (rScreen.Left + ((rScreen.Right - rScreen.Left) / 2) - (oObject.Width / 2))
                oObject.Top = (rScreen.Top + ((rScreen.Bottom - rScreen.Top) / 2) - (oObject.Height / 2))

            Case StartUpPositionConstants.vbStartUpWindowsDefault '0,0
                oObject.Left = rScreen.Left
                oObject.Top = rScreen.Top
        End Select
    End If

    Exit Sub
errsub:
    'vbStartUpWindowsDefault
    oObject.Left = 0
    oObject.Top = 0
End Sub

Private Sub AdjustWindowPositionIntoWorkArea(ByRef oObject As Object, ByRef rScreen As RECT)
'Modified to work for VBA UserForm control which has units measured in points
    On Error GoTo errsub

    oObject.StartUpPosition = vbStartUpManual

    'Adjust window poition as necesssary so that whole window is visible on the screen.
    If TypeOf oObject Is UserForm Then
        'If either left or top coordinate is 0, seems to revert back to windows default settings regardless of what is set here.
        If oObject.Left = 0 Then oObject.Left = 1
        If oObject.Top = 0 Then oObject.Left = 1
    End If

    'Left
    If oObject.Left < rScreen.Left Then oObject.Left = rScreen.Left
    'Top
    If oObject.Top < rScreen.Top Then oObject.Top = rScreen.Top
    'Right
    If oObject.Left + oObject.Width > rScreen.Right Then oObject.Left = rScreen.Right - oObject.Width
    'Bottom
    If oObject.Top + oObject.Height > rScreen.Bottom Then oObject.Top = rScreen.Bottom - oObject.Height

    Exit Sub
errsub:
    'vbStartUpWindowsDefault
    oObject.Left = 0
    oObject.Top = 0
End Sub
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

Private Function GetScreenInfo() As SIZE
'Same as GetDeviceCaps( hdcPrimaryMonitor, HORZRES)
    Dim Screen As SIZE

    Screen.Width = GetSystemMetrics(SM_CXSCREEN)  'in pixels
    Screen.Height = GetSystemMetrics(SM_CYSCREEN) 'in pixels
    GetScreenInfo = Screen
End Function

Private Function PointsPerPixelX(oObject As Object) As Single
'Get UserForm client coordinates and convert points to pixels
    Dim hwndObj As Long
    Dim rRect As RECT

    If TypeOf oObject Is UserForm Then
        hwndObj = FindWindow(C_VBA6_USERFORM_CLASSNAME, oObject.Caption)
    Else
        hwndObj = oObject.hWnd
    End If

    Call GetWindowRect(hwndObj, rRect)
    PointsPerPixelX = oObject.Width / (rRect.Right - rRect.Left) 'Points/Pixel
End Function

Private Function PointsPerPixelY(oObject As Object) As Single
'Get UserForm client coordinates and convert points to pixels
    Dim hwndObj As Long
    Dim rRect As RECT

    If TypeOf oObject Is UserForm Then
        hwndObj = FindWindow(C_VBA6_USERFORM_CLASSNAME, oObject.Caption)
    Else
        hwndObj = oObject.hWnd
    End If

    Call GetWindowRect(hwndObj, rRect)
    PointsPerPixelY = oObject.Height / (rRect.Bottom - rRect.Top)
End Function

