VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCalendar 
   Caption         =   "Calender"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2580
   OleObjectBlob   =   "frmCalendar.frx":0000
End
Attribute VB_Name = "frmCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'File:   frmCalendar
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        1/23/13
'
'Summary:
'**Requires legacy component Mscomct2.ocx which is not always installed and registered.
'Calendar date picker that pops up based on mouse position.
'Implemented with "Microsoft MonthView Control 6.0 (SP4)"
'
'Concept from:
'http://www.yogeshguptaonline.com/2010/01/excel-macros-excel-date-picker.html
'http://www.vbaexpress.com/kb/getarticle.php?kb_id=431
'http://msdn.microsoft.com/en-us/library/aa733656(v=vs.60).aspx
'
'Sample Implementation:
'Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'    On Error GoTo ErrSub
'
'    Dim iCal As frmCalendar
'    Dim rngTarget As Excel.Range
'    Dim retDate As Variant
'
'    Set rngTarget = Me.Range("G12:G14") 'Multiple continguous cells with dates.
'    Set Target = Target(1) 'Only first cell if multi-selection is considered for test.
'
'    If Not Application.Intersect(Target, rngTarget) Is Nothing Then
'        Set iCal = New frmCalendar
'        retDate = iCal.ShowCalendar(CDate(Target.Value))
'        If CDbl(retDate) = 0 Then
'            Target.Value = retDate
'        End If
'    End If
'
'ErrSub:
'    Cancel = True
'    Unload iCal
'    Set iCal = Nothing
'End Sub

Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function DrawMenuBar Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function GetCursorPos Lib "User32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame" 'For VBA UserForm class name

Private Const GWL_STYLE = (-16)
Private Const WS_CAPTION = &HC00000

Private Type POINTAPI
    X As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private WithEvents oMonthView As MonthView 'Ideally would like to late bind (so that reference isn't needed), however can't late bind with events.
Attribute oMonthView.VB_VarHelpID = -1
Private mDate As Date
Private mESC As Boolean

Public Property Get CalendarDate() As Date
    CalendarDate = mDate
End Property

Private Property Let CalendarDate(ByVal calDate As Date)
    mDate = calDate
End Property

Private Sub UserForm_Initialize()
    On Error GoTo errsub
    
    Dim dwStyle As Long
    Dim dwRemove As Long
    Dim hwnd As Long
    Dim p As POINTAPI
    
    If val(Application.Version) >= 9 Then
        hwnd = FindWindow(C_VBA6_USERFORM_CLASSNAME, Me.Caption)
    Else
        hwnd = FindWindow("ThunderXFrame", Me.Caption)
    End If
    
    If hwnd > 0 Then
        dwStyle = GetWindowLong(hwnd, GWL_STYLE)
        dwStyle = dwStyle And Not WS_CAPTION
        Call SetWindowLong(hwnd, GWL_STYLE, dwStyle)
        Call DrawMenuBar(hwnd) 'Redraw menu bar
        Call GetCursorPos(p) 'Position of mouse
        
        Me.Left = p.X * PointsPerPixelX(Me)
'        Me.Left = p.X * Application.ActiveWindow.PointsToScreenPixelsX(p.X)
        
        Me.Top = p.y * PointsPerPixelY(Me)
'        Me.Top = p.Y * Application.ActiveWindow.PointsToScreenPixelsY(Me.Top)
    End If
    
    Set oMonthView = Me.Controls.Add("MSComCtl2.MonthView") 'Added manually as a step toward late binding, however late binding is not implemented as can't late bind with events.

    With oMonthView
'        .Width = 129.75
'        .Height = 118.5
        .TitleBackColor = RGB(160, 160, 160) 'H8000000A
    End With
    Exit Sub
    
errsub:
    Call MsgBox("Unable to load MonthView control. MonthView requires a registered version of Mscomct2.ocx, found in the Windows System or System32 directory.", vbInformation)
End Sub

Private Sub SetRef()
'Dynamically set the reference to Microsoft Windows Common Controls 6.0-2 (SP4)
    On Error Resume Next '< error = reference already set
    ThisWorkbook.VBProject.References.AddFromGuid "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}", 2, 0
End Sub
    
Private Sub oMonthView_DateClick(ByVal DateClicked As Date)
    CalendarDate = oMonthView.Value
    Me.Hide
End Sub

Private Sub cmdButton_Click() 'Used to enable ESC key to close dialog box.
    Me.Hide
    mESC = True
End Sub

Public Function ShowCalendar(Optional InitialDate As Date, Optional MinDate As Date, Optional MaxDate As Date) As Variant 'Date
    If CDbl(InitialDate) > 0 Then 'Check if empty
        oMonthView.Value = InitialDate
    Else
        oMonthView.Value = Date
    End If
    
    If CDbl(MinDate) > 0 Then
        oMonthView.MinDate = MinDate
    End If
    
    If CDbl(MaxDate) > 0 Then
        oMonthView.MaxDate = MaxDate
    End If
    
    Me.Show
    
    If mESC = False Then
        ShowCalendar = CalendarDate
    End If
'    Unload Me
End Function

Public Sub AddInputRangeMessage(rng As Object) 'Excel.Range
'Optionally call to add message to cell that shows when Excel cell is selected.
    With rng.Validation
        .Delete
        Call .Add(xlValidateInputOnly)
        .InputTitle = "Calendar"
        .InputMessage = "Double-click for calendar editor."
        .ShowInput = True
    End With
End Sub

'=======================================================================================================
'Helper Functions
'=======================================================================================================

Private Function PointsPerPixelX(oObject As Object) As Single
'Get UserForm client coordinates and convert points to pixels
    Dim hwndObj As Long
    Dim rRect As RECT
    
    If TypeOf oObject Is UserForm Then
        hwndObj = FindWindow(C_VBA6_USERFORM_CLASSNAME, oObject.Caption)
    Else
        hwndObj = oObject.hwnd
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
        hwndObj = oObject.hwnd
    End If

    Call GetWindowRect(hwndObj, rRect)
    PointsPerPixelY = oObject.Height / (rRect.Bottom - rRect.Top)
End Function
