VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmListSelection 
   Caption         =   "Selection"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "frmListSelection.frx":0000
End
Attribute VB_Name = "frmListSelection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'File:   frmListSelection
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        7/19/2012
'
'Summary:
'Creates a list check box.
'
'Online References:
'Revisions:
'Date     Initials    Description of changes

'Known issues:
'There is known issue in that when double click is used to launch this box, an inadvertant selection can then be made in the list box.
'Modified from source found in Walkenbach Chapter 14 "UserForm Examples" p.466
'Legend seems to remove items as they are hidden, but only to a point.  Can manually remove with mChtSource.Legend.LegendEntries(1).delete

'Notes:
'ListBox1.ListIndex is 0 based
'In order for vertical scroll bar to function correctly interior list box should be 21.25 more than height of window, when interior window completely fills window frame.

'Sample Implementation:
'Dim fSelection As frmListSelection
'Set fSelection = New frmListSelection
'fSelection.AddListItem oPM.RDTGetStreamName(i), i
'...
'fSelection.Show

Private mPreviousCursor As Excel.XlMousePointer 'Long

Private Sub cmdOk_Click()
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    mPreviousCursor = Application.Cursor
    Application.Cursor = xlDefault
End Sub

Private Sub UserForm_Initialize()
'    ListBox1.RowSource = "" 'Tip from Walkenbach: Use to avoid bug in VBA when RowSource is empty and items are to be added.
    Dim oPW As New clsPositionWindow
    Call oPW.ForceWindowIntoWorkArea(Me, vbStartUpCenterParent)
    Set oPW = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then 'Hide rather than closing object.
        Call ToggleListBoxSelections(False)
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    Application.Cursor = mPreviousCursor
End Sub

Public Sub AppendListItem(vListItem As Variant, vTag As Variant) 'colSource As Collection)
    'Add item to listbox
    With ListBox1
        .AddItem vListItem '0,0
        .List(.ListCount - 1, 1) = vTag '0,1
'        .ListIndex = 0 'set to first item in list.
    End With
End Sub

Private Sub cmdAll_Click()
    Call ToggleListBoxSelections(True)
End Sub

Private Sub cmdNone_Click()
    Call ToggleListBoxSelections(False)
End Sub

Private Function ToggleListBoxSelections(Selection As Boolean)
    Dim lSeriesCounter As Long

    For lSeriesCounter = 0 To ListBox1.ListCount - 1
        ListBox1.Selected(lSeriesCounter) = Selection
    Next
End Function

Public Function GetSelectedCount() As Long
    Dim lSeriesCounter As Long

    For lSeriesCounter = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(lSeriesCounter) = True Then '0 based
            GetSelectedCount = GetSelectedCount + 1
        End If
    Next
End Function

Public Function GetSelectedArray() As Variant
    Dim arryRet As Variant
    Dim i As Long
    Dim iCount As Long
    
    iCount = GetSelectedCount
    
    If iCount > 0 Then
        ReDim arryRet(0 To iCount - 1, 0 To 1)
        iCount = 0
        For i = 0 To Me.ListBox1.ListCount - 1
            If Me.ListBox1.Selected(i) = True Then
                arryRet(iCount, 0) = Me.ListBox1.List(i, 0) 'Stream Name
                arryRet(iCount, 1) = Me.ListBox1.List(i, 1) 'Stream ID
                iCount = iCount + 1
            End If
        Next
        GetSelectedArray = arryRet
    End If
End Function
