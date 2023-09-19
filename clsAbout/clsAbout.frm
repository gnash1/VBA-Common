VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clsAbout 
   Caption         =   "About  Product"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7800
   OleObjectBlob   =   "clsAbout.frx":0000
End
Attribute VB_Name = "clsAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

'File:   clsAbout
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 Thepiecemaker.com
'Date:        9/4/12

'Summary:
'Standard Application AboutBox.
'txtDescription embedded inside disabled frameDisable to provide enabled text look while not allowing for text selection.
'http://www.vbcity.com/forums/topic.asp?tid=53291

'Revisions:
'Date     Initials    Description of changes

'Example use:
'Call CustomCommandBar.AddControlButton("About", 487, "ThisWorkbook.ShowAbout", "About")
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

Private m_Application As Object
Private m_ApplicationCursor As Long 'XlMousePointer
Private m_ApplicationEnableEvents As Boolean
Private m_ApplicationScreenUpdating As Boolean

Private Sub imgLarge_Click()
    Call GotoWebSite
End Sub

Private Sub imgSmall_Click()
    Call GotoWebSite
End Sub

Private Sub GotoWebSite()
    On Error GoTo errsub
    
    Dim strLink As String
    
    strLink = "http://www.Thepiecemaker.com"
    
    Select Case m_Application.Name 'Application.Value
        Case "Microsoft Excel"
            Call m_Application.ActiveWorkbook.FollowHyperlink(strLink, , True)
        Case "PowerPoint"
            Call m_Application.ActivePresentation.FollowHyperlink(strLink, , True)
    End Select
    
errsub:
End Sub

Private Sub lblCopyright_Click()
    On Error GoTo errsub
    
    Dim strLink As String
    
    strLink = "mailto:support@Thepiecemaker.com &subject=Support Request"
    
    Select Case m_Application.Name 'Application.Value
        Case "Microsoft Excel"
            Call m_Application.ActiveWorkbook.FollowHyperlink(strLink, , True)
        Case "PowerPoint"
            Call m_Application.ActivePresentation.FollowHyperlink(strLink, , True)
    End Select
    
errsub:
End Sub

Private Sub UserForm_Initialize()
    Dim oPW As New clsPositionWindow
    Dim oActive As Object
    Const Default As Long = -4143
    Set m_Application = Application
    
    Select Case m_Application.Name 'Application.Value
        Case "Microsoft Excel"
            Set oActive = m_Application.ThisWorkbook
            m_ApplicationCursor = m_Application.Cursor
            m_Application.Cursor = Default
            m_ApplicationEnableEvents = m_Application.EnableEvents
            m_Application.EnableEvents = True
            m_ApplicationScreenUpdating = m_Application.ScreenUpdating
            If m_Application.ScreenUpdating = False Then 'Reduce repaint flicker
                m_Application.ScreenUpdating = True
            End If
        Case "Microsoft PowerPoint"
            Set oActive = m_Application.ActivePresentation
        Case "Microsoft Visio"
            Set oActive = m_Application.ActiveDocument
            m_ApplicationEnableEvents = m_Application.EventsEnabled
    End Select
    
    lblCopyright = "Copyright © " & Year(Now()) & " The Piecemaker Corporation" & vbCrLf & _
                                    "Phone: (888) 888-8888" & vbCrLf & _
                                    "Fax: (888) 888-8888" & vbCrLf & _
                                    "Email: support@Thepiecemaker.com"
    lblFileInfo.Caption = vbNullString 'Clear ddefault text so that label doesn't show if not populated.
    
    Call oPW.ForceWindowIntoWorkArea(Me, vbStartUpCenterParent)
    Set oPW = Nothing
End Sub

Private Sub UserForm_Terminate()
    Select Case m_Application.Name 'Application.Value
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

Public Property Let CaptionText(ByVal strCaption As String)
    Me.Caption = strCaption
End Property

Public Property Let DescriptionText(ByVal strDescription As String)
    On Error Resume Next
    txtDescription = strDescription
    txtDescription.SetFocus 'Required for next line to work.
    txtDescription.CurLine = 0 'Can crash (if only one line of text?)
End Property

Public Property Let FileInfoText(ByVal strFileInfo As String)
    lblFileInfo.Caption = strFileInfo
End Property

