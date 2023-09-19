VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomCommandBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File:      clsCustomCommandBar
'Author:    Greg Harward
'Contact:   gharward@gmail.com
'Date:      1/6/16
'
'To Do:

'Summary:
'Functionality for creating and removing customized menus and toolbars.
'Supports a single command bar per application.
'Assumes that commandbar is recreated on open or when reactivated with focus.
'
'Online References:
'FaceIDs Online - http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
'http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
'Create CommandBar with buttons - http://msdn2.microsoft.com/en-us/library/scff9c7c(VS.80).aspx
        
'Revisions:
'Date     Initials    Description of changes

'Sample Impelementaion:
'Excel 2007 Workbook Example Implementation - The following code goes in ThisWorkbook:
'=======================================================================================================
'Option Explicit
'
'Private Sub Workbook_Open()
'    On Error Resume Next
'    Dim CommandBarPopupUID As String
'
'    Set CustomCommandBar = New clsCustomCommandBar
'    With CustomCommandBar
'        Call .AddControlButton("SubroutineToCall", 15795, "ThisWorkbook.SubroutineToCall")
'            Call .AddPopupButton("Excel Loader-Update")
'            Call .AddPopupButtonItem("Format IDs as: Candidate Project Code-ID", 0, "ThisWorkbook.UpdateCandidateProjectCode", vbNullString, False)
'            Call .AddPopupButtonItem("Format IDs as: Candidate Compound Name", 0, "ThisWorkbook.UpdateCandidateProjectName", vbNullString, False)
'        Call .AddControlButton("View Errors", 1090, "ThisWorkbook.ViewErrors")
'        Call .AddControlButton("About", 487, "ThisWorkbook.ShowAbout")
'    End With
'    Application.EnableEvents = True
'End Sub
'
'Private Sub Workbook_BeforeClose(Cancel As Boolean)
'    If Cancel = False Then
''        On Error Resume Next
'        Set CustomCommandBar = Nothing
'    End If
'End Sub
'
'Private Sub Workbook_Activate()
''   Always reload (reload includes internal delete)
'    Call Workbook_Open
'End Sub
'
'Private Sub Workbook_WindowDeactivate(ByVal Wn As Window)
'    Call Workbook_BeforeClose(False)
'End Sub
'
'Private Sub SubroutineToCall()
'    Call MsgBox("Put code here to call when button is pushed.")
'End Sub
'
'Private Sub ShowAbout()
'    Dim oAbout As clsAbout
'    Set oAbout = New clsAbout
'
'    oAbout.CaptionText = "About"
'    oAbout.DescriptionText = "About Utility." & vbCrLf & vbCrLf & _
'                            "<Project information goes here>"
'
'    Call oAbout.Show
'    Set oAbout = Nothing
'End Sub
'
'=======================================================================================================
'Notes:
'=======================================================================================================
'Common FaceIDs: (Some are new in 2007) FaceID range goes to 16,200
'Use Add-In tool found here to see them: http://skp.mvps.org/faceid.htm
'Lightning: 527, 107, 144, 793, 581, 582,346, 640, 786, 793, 1709
'MS Project: 3196
'MS Excel Related: 66, 142, 263, 566, 2105, 2586, 3892, 4165, 16928, 16014-17
'Validate: 166, 329
'DB: 528, 1814, 2071, 2657, 3897, 2109, 2116, 1636, 1631, 1632, 3897
'Others: 2659, 1088-1090, 1716
'Transfer 1785, 1807
'Flag: 1090, 1678, 1885
'Check: 1087, 1664, 1715, 1907, 1997
'Export: 6127, 13917, 15796
'Import: 6127, 15795, 16928
'Erase: 1716, 7502, 2060, 5416, 6049
'Delete: 1019
'Scripts: 521, 588, 629, 1777, 1839, 3463, 3464
'Gantt: 2138, 2147, 2099, 2107, 2108, 2149, 2150, 2168, 2170
'Browse: 1661
'=======================================================================================================
'FaceIDs Online - http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
'Code to Print FaceIDs
'Private Function Item_Open()
'    Dim objCBs
'    Dim objMyCB
'    Dim objButton
'    Dim i
'
'    Set objCBs = Item.GetInspector.CommandBars
'    Set objMyCB = objCBs.Add("Icons")
'    objMyCB.Visible = True
'    For i = 0 To 500
'        Set objButton = objMyCB.Controls.Add
'        objButton.style = 3
'        objButton.Caption = i
'        objButton.FaceId = i
'        objButton.Visible = True
'    Next
'End Function

Private m_CustomCommandBar As Office.CommandBar 'Stores last created Command Bar
Private m_CustomCommandBarPopup As Office.CommandBarPopup 'Stores last created Popup Bar

Const CommandBarID = "Tableau Global Services"

Private Sub Class_Terminate()
    Call DeleteCommandBar
End Sub

Private Sub Class_Initialize()
    If Application.name = "Microsoft Excel" Or Application.name = "Microsoft PowerPoint" Or Application.name = "Microsoft Visio" Then
        Call DeleteCommandBar
        
        If CInt(val(Application.Version)) >= 10 Then '11 = 2003
            Call AddCommandBar
        End If
    End If
End Sub

Public Function Hide()
    m_CustomCommandBar.Visible = False
End Function

Public Function Show()
    m_CustomCommandBar.Visible = True
End Function

Private Function AddCommandBar()
    'Add command bar
    Set m_CustomCommandBar = Application.CommandBars.Add(CommandBarID, msoBarTop, False, True)  'Will crash if commandbar already exists by same name.
    With m_CustomCommandBar
        .Protection = msoBarNoCustomize 'Not customizable
'        .Context = str(visUIObjSetDrawing) & "*"   'Sets where commandbar is to be saved.
    End With
End Function

Public Function AddControlButton(strCaption As String, lButtonFaceID As Long, strCallbackFunction As String, Optional strToolTip As String, Optional bBeginGroup As Boolean = False)
    'Add button control
    Dim CustomCommandBarButton As Office.CommandBarButton
    If Not m_CustomCommandBar Is Nothing Then 'Could be nothing when another commandbar by the same name exists.
        Set CustomCommandBarButton = m_CustomCommandBar.Controls.Add(msoControlButton, , , , True)
        With CustomCommandBarButton
            .BeginGroup = bBeginGroup  'Only true for first button in group which is defined with vertical seperator
            .Caption = strCaption
            .State = msoButtonUp
            If .Caption <> vbNullString Then
                .style = msoButtonIconAndCaption 'msoButtonIconAndCaption
            Else
                .style = msoButtonIcon 'msoButtonIconAndCaption
            End If
            .TooltipText = IIf(strToolTip = vbNullString, strCaption, strToolTip)
            .FaceId = lButtonFaceID
            .OnAction = strCallbackFunction 'Callback function can be function in *.mod or in ThisWorkbook (which is a class)
            .DescriptionText = "ProModel Custom"
            .Tag = vbNullString
            .Parameter = vbNullString 'Like another tag
        End With

        'Show custombar
        If m_CustomCommandBar.Enabled = True Then
            Call Show
        End If
    End If
End Function

Public Function AddPopupButton(strCaption As String, Optional strToolTip As String = vbNullString, Optional bBeginGroup As Boolean = False)
    Dim CommandBarPopup As Office.CommandBarPopup

    'Add button control
    If Not m_CustomCommandBar Is Nothing Then 'Could be nothing when another commandbar by the same name exists.
        Set m_CustomCommandBarPopup = m_CustomCommandBar.Controls.Add(msoControlPopup, , , , True)
        With m_CustomCommandBarPopup
            .BeginGroup = bBeginGroup  'Only true for first button in group which is defined with vertical seperator
            .Caption = strCaption
            .TooltipText = IIf(strToolTip = vbNullString, strCaption, strToolTip)
''            .OnAction = strCallbackFunction 'Callback function can be function in *.mod or in ThisWorkbook (which is a class)
'            .DescriptionText = "Description Text"
            .Tag = vbNullString
            .Parameter = vbNullString 'Like another tag
        End With
        
        'Show
        If m_CustomCommandBar.Enabled = True Then
            Call Show
        End If
    End If
End Function

Public Function AddPopupButtonItem(strCaption As String, lButtonFaceID As Long, strCallbackFunction As String, Optional strToolTip As String = vbNullString, Optional bBeginGroup As Boolean = False)
'Add item to Popup Button
'Pass parameters through callback function example: (http://www.ozgrid.com/forum/showthread.php?t=28365)
'1 param: string
'   strCallbackFunction = "'ThisWorkbook.ChangeWorksheetFocus " & VBA.Chr(34) & wks.CodeName & VBA.Chr(34) & "'"
'2 params: string, number
'   strCallbackFunction = "'ThisWorkbook.ChangeWorksheetFocus " & VBA.Chr(34) & wks.CodeName & VBA.Chr(34) & ", " & wks.Index & "'"
    If Not m_CustomCommandBarPopup Is Nothing Then
        With m_CustomCommandBarPopup.Controls.Add(msoControlButton)
            .BeginGroup = bBeginGroup
            '            .State = msoButtonMixed
            .Caption = strCaption
            .TooltipText = IIf(strToolTip = vbNullString, strCaption, strToolTip)
            .FaceId = lButtonFaceID
'            If lButtonFaceID > 0 Then 'Doesn't seem to make a difference.
'                .style = msoButtonIconAndCaption
'            Else
                .style = msoButtonAutomatic
'            End If
            .OnAction = strCallbackFunction
            .Tag = vbNullString
            .Parameter = vbNullString 'Like another tag
        End With
    End If
End Function

Private Function DeleteCommandBar()
'Delete previous crash relics.
    On Error GoTo errsub
    If m_CustomCommandBar Is Nothing Then
        Set m_CustomCommandBar = Application.CommandBars(CommandBarID) 'Can error if not found
    End If
    
    If Not m_CustomCommandBar Is Nothing Then
        m_CustomCommandBar.Delete 'Can error here is manually deleted from toolbar.
        Set m_CustomCommandBar = Nothing
    End If
errsub:
End Function

Private Function GetCommandBarButtonByTag(Tag As String) As Office.CommandBarButton
    Set GetCommandBarButtonByTag = Application.CommandBars.FindControl(Office.MsoControlType.msoControlButton, , Tag)
End Function

Private Function GetCommandBarButtonByID(ID As Long) As Office.CommandBarButton
    Set GetCommandBarButtonByID = Application.CommandBars.FindControl(Office.MsoControlType.msoControlButton, ID)
End Function

Private Function GetCommandBarButtonByParameter(Parameter As String) As Office.CommandBar
    'Parameters are similar to .Tag  Function is to find by parameter if used.
    Dim CommandBarButton As Office.CommandBarButton
    For Each CommandBarButton In Application.CommandBars
        If StrComp(CommandBarButton.Parameter, Parameter) = 0 Then
            Set GetCommandBarButtonByParameter = CommandBarButton
        End If
    Next
End Function

