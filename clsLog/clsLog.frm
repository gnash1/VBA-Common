VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clsLog 
   Caption         =   "Log"
   ClientHeight    =   11520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15915
   OleObjectBlob   =   "clsLog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
'File:   clsLog
'Author:      Greg Harward
'Date:        1/11/13
'
'Summary:
'Error log class.  Can display errors in window or export to file.  Initial implementation only supports Excel worksheet list.
'Designed to allow for errors to be put into flexible set of collections contained in dynamic array.

'To do:
' - Add back in error logs other than Excel.
' - Need to make window modeless.

'Online References:
'http://www.exceltip.com/st/Log_files_using_VBA_in_Microsoft_Excel/493.html
'Revisions:
'Date     Initials    Description of changes

'Sample Implementation:
'Dim Log As clsLog
'Set Log = New clsLog
'Call Log.Initialize(ExcelFile, shtLog)
'Call Log.Add(1,"Test Error Test", "strRange")
'Call Log.Add(2,"Test Error Test", "strRange2")
'Call Log.Add(3,"Test Error Test", "strRange3")
'Call Log.Add(2,"Test Error Test", "strRange4")
'Call Log.Add(1, "Test Error Test", "strRange5")
'Call Log.Export

Public Enum LogFileTypes
'    Unknown = 0
    TextInWindow = 1
    LogFile
    ExcelFile
    'Excel2003File
    'Excel2007File
End Enum

Public Enum eClassification
    NonCritical = 0
    Critical
    Information
End Enum

Private Enum eData
'    Category = 0
    CategoryText = 0
    SpecificText
    HyperlinkReference
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

Private Enum PathParseMode
    Path
    FileName
    FileExtension
    FileNameWithoutExtension
End Enum

Private m_eFileType As LogFileTypes
Private m_eEntriesExist As Boolean
Private m_colEntries() As New Collection 'Array of collections to hold entries.
Private m_colEntriesBackup() As New Collection 'Array of collections to hold entries.
Private m_oDestination As Object 'Object holding log destination.  Can be worksheet, userform, or text file.
Private m_CriticalCount As Long
Private m_NonCriticalCount As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Public Event UpdateProgress(ByVal PercentProgress As Single, ByVal StatusBarText As String)

Private Property Get LogPath() As String
    LogPath = m_oDestination.GetText
End Property

Public Sub Backup()
    Dim lCategory As Long
    Dim Err As Variant
    
    Erase m_colEntriesBackup 'Erase array
    
    If IsArrayEmpty(m_colEntries) = False Then
        ReDim m_colEntriesBackup(0 To UBound(m_colEntries)) As New Collection
        
        For lCategory = 0 To UBound(m_colEntries)
    '        If Not m_colEntries(lCategory) Is Nothing Then
                For Each Err In m_colEntries(lCategory)
                    Call m_colEntriesBackup(lCategory).Add(Err)
                Next
    '        End If
        Next lCategory
    End If
End Sub

Public Sub Restore()
    Dim lCategory As Long
    Dim Err As Variant
    
    Erase m_colEntries 'Erase array
    
    If IsArrayEmpty(m_colEntriesBackup) = False Then
        ReDim m_colEntries(0 To UBound(m_colEntriesBackup)) As New Collection
        
        For lCategory = 0 To UBound(m_colEntriesBackup)
            For Each Err In m_colEntriesBackup(lCategory)
                Call m_colEntries(lCategory).Add(Err)
            Next
        Next lCategory
    End If
End Sub

Public Function Initialize(eExportType As LogFileTypes, oDestinationFile As Object) As Boolean ', Optional bAppend As Boolean = False)
'Pass in Worksheet object or text as DataObject
    On Error GoTo ErrSub
    
    m_eFileType = eExportType
'    m_bAppend = bAppend
    Select Case eExportType
        Case LogFileTypes.LogFile   'As DataObject
            If DeleteTempFile(oDestinationFile.GetText) = True Then
                Set m_oDestination = oDestinationFile
            End If
        Case LogFileTypes.TextInWindow 'Nothing
            Set m_oDestination = Nothing
        Case LogFileTypes.ExcelFile 'as Excel.Worksheet
            Dim oData As DataObject 'Workaround where in contents of clipbaord are removed in the following lines of code if in the middle of a copy paste operation.
            Set oData = GetClipboard
            Set m_oDestination = oDestinationFile 'worksheet reference.
            '    If bAppend = False Then
            'Clear previous outlining - Needs to be before "ClearErrors" since UsedRange is used.
            Call RemoveOutline(oDestinationFile.UsedRange)
            With oDestinationFile.Outline
                .AutomaticStyles = False
                .SummaryRow = xlAbove
            End With
            '    End If
            Call SetClipboard(oData)
        Case Else
            Call MsgBox("Unimplemented", vbCritical, "Internal Error - " & Me.name)
            Debug.Assert False
    End Select
    
    Call ClearErrors
    
    Initialize = True
ErrSub:
End Function

Public Function AddEntry(lCategory As Long, strCategoryText As String, strSpecificText As String, Optional rngReference As Excel.Range, Optional Classification As eClassification = eClassification.NonCritical)
    'Array of collections
    If IsArrayEmpty(m_colEntries) Then
        ReDim m_colEntries(0 To lCategory) As New Collection
'        m_colEntries(0).Add (strErrorText) 'Used to hold description.
    ElseIf UBound(m_colEntries) < lCategory Then
        ReDim Preserve m_colEntries(0 To lCategory) As New Collection
    End If
    
    If rngReference Is Nothing Then
        Call m_colEntries(lCategory).Add(strCategoryText & "|" & strSpecificText & "|" & vbNullString)
    Else
        Call m_colEntries(lCategory).Add(strCategoryText & "|" & strSpecificText & "|" & rngReference.Address(, , , True))
    End If
    
    Select Case Classification
        Case eClassification.NonCritical
            m_NonCriticalCount = m_NonCriticalCount + 1
        Case eClassification.Critical
            m_CriticalCount = m_CriticalCount + 1
    End Select
    
    m_eEntriesExist = True
'    m_strLogText = m_strLogText & vbCrLf & strText
End Function

Public Sub ClearErrors()
    On Error GoTo ErrSub
    'Clear Objects
    Erase m_colEntries 'Erase array
'    Erase m_colEntriesBackup 'Erase array

    Select Case m_eFileType
        Case LogFileTypes.LogFile
            'Kill previous file
'            If Dir(m_oDestination.GetText) <> vbNullString And m_oDestination.GetText <> vbNullString Then
'                Call Kill(m_oDestination.GetText)
'            End If
        Case LogFileTypes.TextInWindow
            txtLog.Text = vbNullString
        Case LogFileTypes.ExcelFile
            Call RemoveOutline(m_oDestination.UsedRange) 'Do this before clearing contents.
'            Call m_oDestination.UsedRange.Clear 'Contents.  This can be slow.
            Call m_oDestination.UsedRange.EntireRow.Delete
            Call m_oDestination.UsedRange.EntireColumn.Delete
        Case Else
    End Select
ErrSub:
    m_eEntriesExist = False
    m_CriticalCount = 0
    m_NonCriticalCount = 0
End Sub

Public Function EntriesMade() As Boolean
    EntriesMade = m_eEntriesExist
End Function

Public Function CriticalFound() As Boolean
    CriticalFound = (m_CriticalCount > 0)
End Function

Public Function NonCriticalFound() As Boolean
    NonCriticalFound = (m_NonCriticalCount > 0)
End Function

Public Function Export(Optional lReportLimit As Long = 10000) 'eExportType As LogFileTypes, Optional WriteDestination As Variant, Optional bShowScoreboardCount As Boolean)
    On Error GoTo ErrSub
    
    Dim strEntries As String
    Dim vEntry As Variant
    Dim vEntryPieces As Variant
    Dim iEntryCategoryCount As Long
    Dim iEntryCount As Long
    Dim lTotalEntryCount As Long
    Dim eExportType As LogFileTypes
    Dim WriteDestination As Object 'Variant
    Dim rngLogStart As Excel.Range
    Dim rngLogEnd As Excel.Range
    Dim rng As Excel.Range
    Dim wksLog As Excel.Worksheet
    Dim strTextLine As String
    
    'Put limit on output for efficency.
    If lReportLimit > 10000 Then
        lReportLimit = 10000
    End If
                    
    If IsArrayEmpty(m_colEntries) = False Then 'm_eEntriesExist = true
        'Total Errors
        For iEntryCategoryCount = 1 To UBound(m_colEntries)
            lTotalEntryCount = lTotalEntryCount + m_colEntries(iEntryCategoryCount).Count
        Next
        
        'VarType(vPropertyValue)
        eExportType = m_eFileType
        Select Case eExportType
            Case LogFileTypes.LogFile
'                Call MsgBox("Unimplemented", vbCritical, "Internal Error - " & Me.Name)
'                Debug.Assert False
                
                'To write the data dimension and fill array A, then ....
                Dim FileNumber As Integer
                Dim yFileBuffer As Variant

                FileNumber = FreeFile
                
                Open LogPath For Output As #FileNumber 'Create a new file and record if file does not exist
                
                'Write out errors
            'Header
                Print #FileNumber, "ProModel Corporation Log File"
                Print #FileNumber, Now()
                Print #FileNumber, String(100, "=")

                'Write out errors
                For iEntryCategoryCount = 1 To UBound(m_colEntries)
                    
                    For iEntryCount = 1 To m_colEntries(iEntryCategoryCount).Count 'First one holds header.
                        RaiseEvent UpdateProgress((iEntryCount / lTotalEntryCount) * 100, "Building Error Log: " & iEntryCount & "/" & lTotalEntryCount)
                        vEntryPieces = Split(m_colEntries(iEntryCategoryCount).Item(iEntryCount), "|")
'                        If iEntryCount = 1 Then 'Section header
'                            strTextLine = vEntryPieces(eData.CategoryText)
'                            strTextLine = strTextLine & vbTab & m_colEntries(iEntryCategoryCount).Count
'                        End If

                        Print #FileNumber, vEntryPieces(eData.CategoryText)
                        Print #FileNumber, vEntryPieces(eData.SpecificText)
'                        strTextLine = vEntryPieces(eData.HyperlinkReference)
                        
'                        Print #FileNumber, strTextLine  'Write information
'                        Print #FileNumber, vEntryPieces(eData.SpecificText)
                        Print #FileNumber, vbCrLf
                    Next
                Next
                
                Close #FileNumber
                
            Case LogFileTypes.TextInWindow
                Call MsgBox("Unimplemented", vbCritical, "Internal Error - " & Me.name)
                Debug.Assert False
                
                For iEntryCategoryCount = 1 To UBound(m_colEntries)
                    For iEntryCount = 1 To m_colEntries(iEntryCategoryCount).Count 'First one holds header.
                        RaiseEvent UpdateProgress((iEntryCount / lTotalEntryCount) * 100, "Building Error Log: " & iEntryCount & "/" & lTotalEntryCount)
                        vEntryPieces = Split(m_colEntries(iEntryCategoryCount).Item(iEntryCount), "|")
                        If iEntryCount = 1 Then
                            'Header
                            rng.Value = vEntryPieces(eData.CategoryText)
                            rng.Offset(0, 1).Value = m_colEntries(iEntryCategoryCount).Count
                            
                            'Text format
                            With rng.Parent.Range(rng, rng.Offset(0, 1))
                                .Font.Bold = True
'                                    .Font.Underline = xlUnderlineStyleSingle
                                .HorizontalAlignment = Excel.Constants.xlLeft
                            End With
                            Set rng = rng.Offset(1)
                            Set rngLogStart = rng
                        End If
                        
                        'Reference hyperlinks
                        rng.Offset(0, 1).Value = vEntryPieces(eData.SpecificText)
'                            Call AddErrorHyperlinkByAddressText(rng.Offset(0, 1), CStr(vEntryPieces(eErrorData.ErrorHyperlinkReference)))
                        Call AddErrorHyperlinkByAddressText(rng.Offset(0, 1), CStr(vEntryPieces(eData.HyperlinkReference)))
                        
                        Set rng = rng.Offset(1)
                        If rng.Row = lReportLimit Then 'rng.Parent.Rows.Count Then 'Stop if error exceeeds the number of rows that Excel can hold.
                            Exit For
                        End If
                    Next
                    
'                        Call RemoveOutline(wksLog.UsedRange)
                    If Not rngLogStart Is Nothing Then
                        Set rngLogEnd = rng.Offset(-1)
                        Call wksLog.Range(rngLogStart, rngLogEnd).Rows.EntireRow.Group
                    End If
                    Set rngLogStart = Nothing 'Used to determine if new category contains any values to group.
                Next
                    
    '            'Header
    '            m_strLogHeader = "ProModel Corporation - Log File" & vbCrLf & Now()
    '            m_strLogText = m_strLogHeader
    '            For iCount = 0 To UBound(m_colEntries)
    '                vEntryPieces = Split(vEntry, "|")
    '        '        strEntries = strEntries & vEntry
    '
    '                oError.strMessage = vEntryPieces(0)
    '                oError.strReference = vEntryPieces(1)
    '                oError.lType = vEntryPieces(2)
    '            Next
'                txtLog.Text = m_strLogText
            Case LogFileTypes.ExcelFile
                'Log error message.
                
                Set WriteDestination = m_oDestination
                
                'Excel.Constants
                If varType(WriteDestination) = vbObject Then
                    Set wksLog = WriteDestination
                    
                    'Total Summary
                    Set rng = wksLog.Range("$A$1")
    '                Set rng = GetLastUsedCellInColumn(wksLog.Range("$A$1")).Offset(1, 0) '.Row + 1
                    
                    rng.Value = "Total Error Count"
                    Call InsertComment(rng, "Time Stamp: " & Format(Now, "dd-mm-yy hh:mm:ss"))
                    
                    If lTotalEntryCount >= lReportLimit Then
                        rng.Offset(0, 1).Value = "Truncated error report showing: " & lReportLimit & "/" & lTotalEntryCount
                    Else
                        rng.Offset(0, 1).Value = lTotalEntryCount
                    End If
                    
                    'Text format
                    With rng.Parent.Range(rng, rng.Offset(0, 1))
                        .Font.Bold = True
                        .Font.Underline = xlUnderlineStyleSingle
                        .HorizontalAlignment = Excel.Constants.xlLeft
                    End With
                    Set rng = rng.Offset(1)
                   
                    'Write out errors
                    For iEntryCategoryCount = 1 To UBound(m_colEntries)
                        For iEntryCount = 1 To m_colEntries(iEntryCategoryCount).Count 'First one holds header.
                            RaiseEvent UpdateProgress((iEntryCount / lTotalEntryCount) * 100, "Building Error Log: " & iEntryCount & "/" & lTotalEntryCount)
                            vEntryPieces = Split(m_colEntries(iEntryCategoryCount).Item(iEntryCount), "|")
                            If iEntryCount = 1 Then
                                'Header
                                rng.Value = vEntryPieces(eData.CategoryText)
                                rng.Offset(0, 1).Value = m_colEntries(iEntryCategoryCount).Count
                                
                                'Text format
                                With rng.Parent.Range(rng, rng.Offset(0, 1))
                                    .Font.Bold = True
'                                    .Font.Underline = xlUnderlineStyleSingle
                                    .HorizontalAlignment = Excel.Constants.xlLeft
                                End With
                                Set rng = rng.Offset(1)
                                Set rngLogStart = rng
                            End If
                            
                            'Reference hyperlinks
                            rng.Offset(0, 1).Value = vEntryPieces(eData.SpecificText)
'                            Call AddErrorHyperlinkByAddressText(rng.Offset(0, 1), CStr(vEntryPieces(eErrorData.ErrorHyperlinkReference)))
                            Call AddErrorHyperlinkByAddressText(rng.Offset(0, 1), CStr(vEntryPieces(eData.HyperlinkReference)))
                            
                            Set rng = rng.Offset(1)
                            If rng.Row = lReportLimit Then 'rng.Parent.Rows.Count Then 'Stop if error exceeeds the number of rows that Excel can hold.
                                Exit For
                            End If
                        Next
                        
'                        Call RemoveOutline(wksLog.UsedRange)
                        If Not rngLogStart Is Nothing Then
                            Set rngLogEnd = rng.Offset(-1)
                            Call wksLog.Range(rngLogStart, rngLogEnd).Rows.EntireRow.Group
                        End If
                        Set rngLogStart = Nothing 'Used to determine if new category contains any values to group.
                    Next

                    Call wksLog.Outline.ShowLevels(1)
                    
                    'Freeze rows/columns
                    Set rng = Selection
                    wksLog.Select
                    With wksLog.Parent.Windows(1)
                    'Freeze rows/columns based on cell.
                        .FreezePanes = False
                        .SplitRow = 1
                        .SplitColumn = 0
                        wksLog.Range("A1").Select
                        .FreezePanes = True
                    'Freeze first column.
        '                .FreezePanes = False
        '                .SplitColumn = rngGridAnchor.Column - 1
        '                .SplitRow = 0
        '                .FreezePanes = True
                    End With
                    rng.Parent.Select 'Put focus back
                    
                    wksLog.Columns("B:B").WrapText = False
                    wksLog.UsedRange.NumberFormat = "General"
                    wksLog.UsedRange.EntireColumn.AutoFit
                End If
            Case Else
                Debug.Print "Must call 'Initialize()' afer object creation"
'                Call MsgBox("Must call 'Initialize()' afer object creation", vbCritical, Me.Name & " - Internal Error")
                Debug.Assert False
        End Select
    End If
    
ErrSub:
    Close #FileNumber
End Function

Private Sub AddErrorHyperlinkByAddressText(rngAnchor As Excel.Range, strLink As String)
'Add hyperlink by simplified set of text rather than rng.Address(,,,true) so that it will automatically update when worksheet is saved as a new name.
    If strLink <> vbNullString Then
        Dim strTemp As String
        strTemp = "'" & Mid(strLink, InStr(1, strLink, "]") + 1)
        strTemp = "'" & Replace(strTemp, "'", vbNullString)
        strTemp = Replace(strTemp, "!$", "'!")
        strTemp = Replace(strTemp, "$", vbNullString)

        Call rngAnchor.Parent.Hyperlinks.Add(rngAnchor, vbNullString, strTemp, "Address: " & strTemp) ', rngLink.value)
    'Call shtLog.Hyperlinks.Add(rngError.Offset(0, 1), "", rng.Address(, , , True), "Address: " & rng.Address(, , , True))
    End If
End Sub

Public Sub Display()
'    On Error Resume Next
    Dim retVal As Long
    
'    Call Export

    Select Case m_eFileType
        Case LogFileTypes.LogFile
            'http://support.microsoft.com/kb/238245
            'http://www.vbaccelerator.com/codelib/shell/shellex.htm
            retVal = ShellExecute(GetDesktopWindow(), "open", LogPath, "", "", SW_SHOWNORMAL) 'Better than VB "Shell" as it allows launching with file path.
        Case LogFileTypes.TextInWindow
            Call Me.Show(vbModal)
        Case LogFileTypes.ExcelFile
            m_oDestination.Activate 'Set focus to Excel worksheet.
        Case Else
    End Select
    
'    On Error GoTo 0
End Sub

'Public Sub Hide()
'    Me.Hide
'End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''Helper Functions
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Private Function CloseFile(strFilePath As String, Optional SaveBeforeClose As Boolean = False) As Boolean
'    On Error GoTo ErrSub
'    Dim FilePtr As Object
'
'    If IsFileOpen(strFilePath) = True Then
'        Set FilePtr = GetObject(strFilePath)
''        Application.DisplayAlerts = False   ' TURN OFF EXCEL DISPLAY
'        If SaveBeforeClose = True Then
'            FilePtr.Save
'        End If
'        FilePtr.Close
'    End If
'    CloseFile = True
'ErrSub:
'    Set FilePtr = Nothing
''    Application.DisplayAlerts = True
'End Function

Private Function IsFileOpen(strFullFilePath As String) As Boolean
'http://www.xcelfiles.com/IsFileOpen.html
'Check if File is Open.  Does not correctly report for text files.
    Dim hdlFile As Long

    'Error is generated if you try opening a File for ReadWrite lock >> MUST BE OPEN!
    On Error GoTo ErrSub:
    hdlFile = FreeFile
    Open strFullFilePath For Random Access Read Write Lock Read Write As hdlFile
    IsFileOpen = False
    Close hdlFile
    Exit Function
    
ErrSub: 'Someone has file open
    IsFileOpen = True
    Close hdlFile
End Function

Private Function IsArrayEmpty(checkArray As Variant) As Boolean
'Used to tell is array is empty. Uninitialized arrays or arrays that are cleared with Erase() return true.
    On Error GoTo emptyError 'Sometimes errors when testing if = -1 if empty.
    
    If -1 = UBound(checkArray) Then
        GoTo emptyError
    End If

    Exit Function
emptyError:
    Err.Clear
    IsArrayEmpty = True
End Function

Private Sub SetWorksheetsProtection(wkb As Excel.Workbook, bToggle As Boolean)
'Sets protection on: UserInterfaceOnly
    Dim wks As Excel.Worksheet
    For Each wks In wkb.Worksheets
        If bToggle = True Then
            Call wks.Protect(, , , , True)
        Else
            wks.Unprotect
        End If
    Next
End Sub

Private Function GetClipboard() As DataObject
    On Error GoTo ErrSub
    
    Dim oData As DataObject
    
    Set oData = New DataObject
    Call oData.GetFromClipboard
    Set GetClipboard = oData
ErrSub:
End Function

Private Function SetClipboard(oData As DataObject) As Boolean
    On Error GoTo ErrSub
    
    oData.PutInClipboard
    SetClipboard = True

ErrSub:
End Function

Private Function DeleteTempFile(strFullPath As String, Optional strFileExtension As String = "tmp") As Boolean
'Try to kill previous log file.
    On Error GoTo ErrSub

    If Dir(strFullPath) <> vbNullString And strFullPath <> vbNullString Then         'If file already exists
        Call Kill(strFullPath)
    End If
        
    DeleteTempFile = True
ErrSub:
End Function

Private Function CreateFilePath(strFilePath As String) As String
'Only creates one folder level.

    Dim strFolderPath As String

    If strFilePath <> vbNullString Then
        strFolderPath = ParsePath(strFilePath, Path)
        If VBA.Dir(strFolderPath, vbDirectory) = vbNullString Then
            Call MkDir(strFolderPath)
        End If
    End If

    CreateFilePath = strFolderPath
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

    On Error GoTo ErrSub

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
ErrSub:
    Set FSO = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''UserForm Code
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserForm_Initialize()   'Private Sub Class_Initialize()
'    Set m_colErrors = New Collection
    Dim oPW As New clsPositionWindow
    
    Me.Hide
    'Set default position
    Call oPW.ForceWindowIntoWorkArea(Me)
'    Call Initialize(Unknown)

    Set oPW = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then   ' user clicked the X button
        'Cancel unloading the form object
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()
    'Write out results
'    Call Export
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Outline Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function LoaderApplyOutlineRowLevels()
    Call ApplyOutlineRowLevels(ActiveSheet.Columns(1))
End Function

Private Function ApplyOutlineRowLevels(rngLevels As Excel.Range)
    'Removes all outline levels then applies new levels to rows of entire sheet based on values in rngLevels.
    'Levels can be appled with keyboard Shift+ Alt + right arrow | left arrow.
    'Outline levels are 1 based.
    Dim rng As Excel.Range
    
    'Remove previous subtotal
    Call RemoveOutline(rngLevels.Parent)
    
    Set rngLevels = Intersect(rngLevels(1).EntireColumn, rngLevels.Parent.UsedRange)
    For Each rng In rngLevels
        If rng.Value > 0 And IsNumeric(rng.Value) Then 'could be summary row.
            rng.EntireRow.OutlineLevel = val(rng.Value)
        End If
    Next
End Function

Private Function RemoveOutline(rngRemove As Excel.Range)
    'Removes outline levels from worksheet given range.
    On Error Resume Next
    'wks.UsedRange.RemoveSubtotal
    rngRemove.RemoveSubtotal 'Can error if empty
    Err.Clear
End Function

Private Sub InsertComment(rng As Excel.Range, strComment As String)
    rng.ClearComments
    rng.AddComment (Trim(strComment))
    rng.Comment.shape.TextFrame.AutoSize = True
End Sub
