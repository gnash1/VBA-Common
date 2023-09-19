VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCustomCommandBarEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File:   clsCustomCommandBarEvents
'Author:      Greg Harward
'Contact:     gregh@promodel.com
'Copyright © 2012 ProModel Corporation
'Date:        3/30/12
'
'To Do:

'Summary:
'Functionality for creating and removing customized menus and toolbars.
'Fires an Event to notify caller when button is clicked, caller to handle events based on passed in ButtonID.
'Implementation relies on the fact that the same event handler will be fired for items with the same .ID and .Tag.
'Implemented instead of OnAction() which was returning errors when used with Visio.
'Supports a single command bar per application.
'Assumes that commandbar is recreated on open or when reactivated with focus.
'
'Online References:
'FaceIDs Online - http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
'http://www.outlookexchange.com/articles/toddwalker/BuiltInOLKIcons.asp
'Create CommandBar with buttons - http://msdn2.microsoft.com/en-us/library/scff9c7c(VS.80).aspx
        
'Sample Impelementaion:
'Excel 2007 Workbook Example Implementation - The following code goes in ThisWorkbook:
'=======================================================================================================
'Option Explicit
'
'Private WithEvents CustomCommandBar As clsCustomCommandBarEvents
'
'Private Sub Workbook_Open()
'    Dim ComboBox As Office.CommandBarComboBox
'
'    Set CustomCommandBar = New clsCustomCommandBarEvents
'    With CustomCommandBar
'        Call .AddControlButton("SubroutineToCall", 15795)
'        Call .AddPopupButton("Excel Loader-Update")
'            Call .AddPopupButtonItem("Format IDs as: Candidate Project Code-ID", 0)
'            Call .AddPopupButtonItem("Format IDs as: Candidate Compound Name", 0)
'
'        Set ComboBox = .AddDropDownButton("Combo Test", Array("AS IS", "TO BE"))
''        ComboBox.Text 'To get text of selected combo item
'        Call .AddControlButton("View Errors", 1090)
'        Call .AddControlButton("About", 487)
'        Call .Show
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
'Private Sub CustomCommandBar_ButtonClick(ButtonID As String, Item As Long)
'    Select Case ButtonID
'        Case "Page Selection"
'            Select Case Item
'                Case ASIS
'                    Call ChangeHyperlinkTargetPage(eBoeingPage.ASIS)
'                Case TOBE
'                    Call ChangeHyperlinkTargetPage(eBoeingPage.TOBE)
'            End Select
'        Case "About"
'            Call ShowAbout
'    End Select
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
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32.dll" (rclsid As GUID) As Long

Private m_CustomCommandBar As Office.CommandBar 'Stores last created Command Bar

Private m_CustomCommandBarPopup As Office.CommandBarPopup 'Stores last created Popup Bar
Private m_CustomCommandBarComboBox As Office.CommandBarComboBox 'Stores last created Combo Box

Private m_CustomCommandBarButtonID As String

Private WithEvents CustomCommandBarButton As Office.CommandBarButton
Attribute CustomCommandBarButton.VB_VarHelpID = -1
Private WithEvents CustomCommandBarComboBox As Office.CommandBarComboBox
Attribute CustomCommandBarComboBox.VB_VarHelpID = -1

Public Event ButtonClick(ButtonID As String, Item As Long) 'Item is name of selected item in ComboBox.
'Public Event ComboClick(UID As String)

Const CommandBarID = "ProModel"

Private Sub Class_Terminate()
    Call DeleteCommandBar
End Sub

Private Sub Class_Initialize()
    If Application.Name = "Microsoft Excel" Or Application.Name = "Microsoft PowerPoint" Or Application.Name = "Microsoft Visio" Then
        Call DeleteCommandBar
        m_CustomCommandBarButtonID = GetGUIDString()
        If CInt(val(Application.Version)) >= 10 Then '11 = 2003
            Call AddCommandBar
        End If
    End If
End Sub

Public Function Hide()
    m_CustomCommandBar.Visible = False
End Function

Public Function Show()
'    If m_CustomCommandBar.Enabled = True Then
        m_CustomCommandBar.Visible = True
'    End If
End Function

Private Function AddCommandBar()
    'Add command bar
    Set m_CustomCommandBar = Application.CommandBars.Add(CommandBarID, msoBarTop, False, True)  'Will crash if commandbar already exists by same name.
    With m_CustomCommandBar
        .Protection = msoBarNoCustomize 'Not customizable
'        .Context = str(visUIObjSetDrawing) & "*"   'Sets where commandbar is to be saved.
    End With
End Function

'Public Function AddControlButton(strCaption As String, lButtonFaceID As Long, Optional strToolTip As String, Optional bBeginGroup As Boolean = False)
Public Function AddControlButton(strCaption As String, lButtonFaceID As Long, Optional UID As String, Optional strToolTip As String, Optional bBeginGroup As Boolean = False)
    'Add button control
'    Dim CustomCommandBarButton As Office.CommandBarButton
    If Not m_CustomCommandBar Is Nothing Then 'Could be nothing when another commandbar by the same name exists.
'        Set CustomCommandBarButton = m_CustomCommandBar.Controls.Add(msoControlButton, , , , True)
'        With CustomCommandBarButton
        With m_CustomCommandBar.Controls.Add(msoControlButton, , , , True)
            If CustomCommandBarButton Is Nothing Then 'Only need to set it on the first instance, and it will fire for all created buttons with same ID and Tag.
                Set CustomCommandBarButton = .Control
            End If

            .BeginGroup = bBeginGroup  'Only true for first button in group which is defined with vertical seperator
            .Caption = strCaption
            .State = msoButtonUp
            If .Caption <> vbNullString Then
                .style = msoButtonIconAndCaption 'msoButtonIconAndCaption
            Else
                .style = msoButtonIcon 'msoButtonIconAndCaption
            End If
            .TooltipText = IIf(strToolTip = vbNullString, strCaption, strToolTip)
            .FaceID = lButtonFaceID
'            .OnAction = strCallbackFunction 'Callback function can be function in *.mod or in ThisWorkbook (which is a class)
            .DescriptionText = "ProModel Custom"
            .Tag = m_CustomCommandBarButtonID 'vbNullString
            .Parameter = IIf(UID = vbNullString, strCaption, UID) 'Functions like tag
        End With
    End If
End Function

Public Function AddPopupButton(strCaption As String, Optional UID As String, Optional strToolTip As String = vbNullString, Optional bBeginGroup As Boolean = False)
    'Event is fired when item is clicked and not button itself.
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
            .Tag = m_CustomCommandBarButtonID 'vbNullString
            .Parameter = IIf(UID = vbNullString, strCaption, UID) 'Functions like tag
        End With
    End If
End Function

Public Function AddPopupButtonItem(strCaption As String, lButtonFaceID As Long, Optional UID As String, Optional strToolTip As String = vbNullString, Optional bBeginGroup As Boolean = False)
'Public Function AddPopupButtonItem(strCaption As String, lButtonFaceID As Long, strCallbackFunction As String, Optional strToolTip As String = vbNullString, Optional bBeginGroup As Boolean = False)
'Add item to Popup Button
'Pass parameters through callback function example: (http://www.ozgrid.com/forum/showthread.php?t=28365)
'1 param: string
'   strCallbackFunction = "'ThisWorkbook.ChangeWorksheetFocus " & VBA.Chr(34) & wks.CodeName & VBA.Chr(34) & "'"
'2 params: string, number
'   strCallbackFunction = "'ThisWorkbook.ChangeWorksheetFocus " & VBA.Chr(34) & wks.CodeName & VBA.Chr(34) & ", " & wks.Index & "'"
    If Not m_CustomCommandBarPopup Is Nothing Then
        With m_CustomCommandBarPopup.Controls.Add(Office.MsoControlType.msoControlButton)
            .BeginGroup = bBeginGroup
            '            .State = msoButtonMixed
            .Caption = strCaption
            .TooltipText = IIf(strToolTip = vbNullString, strCaption, strToolTip)
            .FaceID = lButtonFaceID
'            If lButtonFaceID > 0 Then 'Doesn't seem to make a difference.
'                .style = msoButtonIconAndCaption
'            Else
                .style = msoButtonAutomatic
'            End If
'            .OnAction = strCallbackFunction
            .Tag = m_CustomCommandBarButtonID 'vbNullString
            .Parameter = IIf(UID = vbNullString, strCaption, UID) 'Functions like tag
        End With
    End If
End Function

Public Function AddDropDownButton(strCaption As String, Items As Variant, Optional UID As String = vbNullString, Optional ListIndex As Long = 1, Optional strToolTip As String = vbNullString, Optional bBeginGroup As Boolean = False) As Office.CommandBarComboBox
'Fires ButtonClick event for default item selection.
    Dim i As Long
    
    If Not m_CustomCommandBar Is Nothing Then 'Could be nothing when another commandbar by the same name exists.
        With m_CustomCommandBar.Controls.Add(Office.MsoControlType.msoControlDropdown)
'            Static CustomCommandBarComboBox
'            If CustomCommandBarComboBox Is Nothing Then
                Set CustomCommandBarComboBox = .Control
'            End If
            
            .BeginGroup = bBeginGroup  'Only true for first button in group which is defined with vertical seperator
            .Caption = strCaption
            .TooltipText = IIf(strToolTip = vbNullString, strCaption, strToolTip)
''            .OnAction = strCallbackFunction 'Callback function can be function in *.mod or in ThisWorkbook (which is a class)
'            .DescriptionText = "Description Text"
            .Tag = m_CustomCommandBarButtonID 'vbNullString
            .Parameter = IIf(UID = vbNullString, strCaption, UID) 'Functions like tag
            
            For i = 0 To UBound(Items) 'Add items
                .AddItem CStr(Items(i))
            Next
            
            If .ListCount > 0 Then 'Shouldn't be empty
                .ListIndex = ListIndex '0 is default of no selection.
                RaiseEvent ButtonClick(.Parameter, .ListIndex) '.List(.ListIndex))
            End If

            Set AddDropDownButton = .Control
        End With
    End If
End Function

Private Function DeleteCommandBar()
'Delete previous crash relics.
    On Error GoTo errsub
    
    If m_CustomCommandBar Is Nothing Then
'        Set m_CustomCommandBar = Application.CommandBars.FindControl(, , m_CustomCommandBarButtonID).Parent
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

Private Sub CustomCommandBarButton_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
'Will fire for any button where the .ID and .Tag properties are the same.
'Fires event for button push. Called to handle using Tab and Parameter properties of CommandBarButton.
'Set CancelDefault = true to disable standard Office functionality for a button.
    If Ctrl.ID = 1 Then ' 1 = Custom 'And Ctrl.Tag = vbNullString Then
'        Call CallByName(ThisWorkbook, "ShowAbout", VbMethod)
        RaiseEvent ButtonClick(Ctrl.Parameter, 0)
    End If
End Sub

Private Sub CustomCommandBarComboBox_Change(ByVal Ctrl As Office.CommandBarComboBox)
'Will fire for any button where the .ID and .Tag properties are the same.
'Fires event for button push. Called to handle using Tab and Parameter properties of CommandBarButton.
'Set CancelDefault = true to disable standard Office functionality for a button.
    If Ctrl.ID = 1 Then ' And Ctrl.Tag = vbNullString Then
        RaiseEvent ButtonClick(Ctrl.Parameter, Ctrl.ListIndex) 'Ctrl.List(Ctrl.ListIndex))
    End If
End Sub

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
