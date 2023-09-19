VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clsExcel 
   Caption         =   "Excel Workbook Selection"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6675
   OleObjectBlob   =   "clsExcel.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clsExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'File:   clsExcel
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        9/21/12

'Summary:
'Created to handle the acquire and release of the Excel object.
'Make sure to check the return object is Not Nothing when using.
'If existing instance of Excel is open when called, that instance is used, else a new instance is created.
'If file path is valid then file is then opened in application instance acquired.
'If file is already opened and on remote application instance, prompts to close the file first.
'This could be replaced with ability to access the remote instance, however using this method of access is very slow so prompting instead.
'If adding this ability would then also need to check on workbook close if app should also be closed.
'Code loses object handle (m_wkExcelWorkbook) is workbook is made visible and user interacts with the workbook.  Object is not nothing, but is also not filled.

'Addresses Microsoft bug described here:
'PRB: Releasing Object Variable Does Not Close Microsoft Excel
'http://support.microsoft.com/kb/132535/EN-US
'http://msdn.microsoft.com/archive/default.asp?url=/archive/en-us/dnaraccessdev/html/ODC_MicrosoftAccessOLEAutomation.asp
'http://msdn2.microsoft.com/en-us/library/e9waz863(VS.71).aspx

'Example Use:
'Private Sub BuildPortfolioBuilderTemplate()
'    Dim wkb As Excel.Workbook
'    Dim strExcelSourceFile As String
'    Dim ExcelWrapper As clsExcel
'    Set ExcelWrapper = New clsExcel
'
'    Set wkb = ExcelWrapper.GetExcelWkbObject(Me.Application, strExcelSourceFile, True)
'    If Not wkb Is Nothing Then
'        '<Code Here>
'        Call ExcelWrapper.CloseExcelWkbObject
'    End If
'End Sub

'Notes:
'Using GetObject(<filepath>) can cause workbook to be saved with invisible worksheets.  Use CreateObject instead making sure to close Application and Workbooks followed by setting objects to nothing.
'Application.UserControl = True may help
'Set to True, Excel will not close when obejct is freed (set to Nothing). Rod Gill p.268 & http://msdn2.microsoft.com/en-us/library/aa814561(VS.85).aspx 'Makes file behave as if opened by user. Bovey p.814
'Problem is if file is already open.  If in existing application or new application instance this is not a problem.
'If open in another instance of Excel then could open that instance, however this is slow.  For speed, the best is to be in the same instance.
'Currently this file is opened, but as read only.
'
'wkb.RunAutoMacros xlAutoOpen - calling runs Auto_Open procedure that is not called when opening object via automation.
'
'Includes code to manually kill Excel instances via WMI scripts , however not currently used.
'Found instances of using when file accessed via another Excel instance was saved in an invisible state.  Check for this using:
'   Application.Windows(ThisWorkbook.Name).Visible = True
'Could always close object if: m_HostApplication.Workbooks.Count = 0
'Can also get Excel Sheet objects using:
'        Set m_wkExcelWorkbook = CreateObject("EXCEL.SHEET")
'        'http://msdn2.microsoft.com/en-us/library/aa814561(VS.85).aspx
'Did not include code for restricting screen updateing of changing curors, though could have using:
'    m_wkExcelWorkbook.Application.ScreenUpdating = False
'    m_wkExcelWorkbook.Application.Cursor = xlWait

Private Const WM_USER = 1024

Public Enum eWorkbookReferenceType
    WorksheetName = 0
    WorksheetCodeName = 1
End Enum

'Use SendMessageTimeout() instead of SendMessage() so that application doesn't hang waiting for response if destination window is unresponsive.
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long

Private m_ExcelWasAlreadyOpen As Boolean
Private m_SaveOnTerminate As Boolean
Private m_LeaveFileOpenOverride As Boolean

Private m_HostApplication As Object 'Application calling for Excel instance.
Private m_ApplicationInstance As Object 'Excel.Application
Private m_wkExcelWorkbook As Object 'Excel.Workbook

Private m_ExcelWorksheetName As String
Private m_colWorkbooks As Collection
Private m_Automation As Office.MsoAutomationSecurity
Private m_ExcelInstanceVisible As Boolean

'Private Const xlWait = 2
'Private Const xlDefault = &HFFFFEFD1    '4143

Public Property Get ApplicationInstance() As Object 'Excel.Application
    Set ApplicationInstance = m_ApplicationInstance
End Property

Private Property Set ApplicationInstance(ByRef Instance As Object) 'Excel.Application
    Set m_ApplicationInstance = Instance
End Property

Public Property Get WorkbookInstance() As Object 'Excel.Application
    Set WorkbookInstance = m_wkExcelWorkbook
End Property

Private Property Set WorkbookInstance(ByRef Workbook As Object) 'Excel.Application
    Set m_wkExcelWorkbook = Workbook
End Property

Private Property Get HostApplication() As Object 'Application
    Set HostApplication = m_HostApplication
End Property

Private Property Set HostApplication(ByRef ApplicationInstance As Object)   'Application)
    Set m_HostApplication = ApplicationInstance
End Property

Public Property Get SaveOnTerminate() As Boolean
'Get property state.
    SaveOnTerminate = m_SaveOnTerminate
End Property

Public Property Let SaveOnTerminate(bSave As Boolean)
'Can be used to cancel save if errors are detected during operation.
    m_SaveOnTerminate = bSave
End Property

'Public Function GetObjectExcelUsingCustomProperties(CallingApplication As Application, strCustomPropertyName As String, Optional bSaveOnTerminate As Boolean = True, Optional bAllowValidWorkbookSelection As Boolean = True, Optional bApplyLeaveFileOpenOverride As Boolean = False) As Object 'Excel.Worksheet
'    Dim strExcelImportFile As String
'    Dim strExcelImportWorksheet As String
'
'    'Get default Excel file to use.
'    strExcelImportFile = GetWorkbookProperty(ThisWorkbook, strCustomPropertyName & "Host", True)
'    strExcelImportWorksheet = GetWorkbookProperty(ThisWorkbook, strCustomPropertyName & "Table", True)
'
'    Set GetObjectExcelUsingCustomProperties = GetObjectExcel(CallingApplication, strExcelImportFile, strExcelImportWorksheet, WorksheetName, bSaveOnTerminate, bAllowValidWorkbookSelection, bApplyLeaveFileOpenOverride)
'
'    If Not GetObjectExcelUsingCustomProperties Is Nothing Then
'        Call SetWorkbookProperty(ThisWorkbook, strCustomPropertyName & "Host", strExcelImportFile, True)
'        Call SetWorkbookProperty(ThisWorkbook, strCustomPropertyName & "Table", strExcelImportWorksheet, True)
'    End If
'End Function

Public Function GetExcelWkbObject(CallingApplication As Application, ByRef strWorkbookFilePath As String, Optional bPromptToUseFile As Boolean = False, Optional bSaveOnTerminate As Boolean = False, Optional bApplyLeaveFileOpenOverride As Boolean = False, Optional strPromptDialogTitle As String = "File Path", Optional bSuppressVBA = True, Optional bCreateNewWorkbookInstance = False) As Object 'Excel.Workbook
'Public Function GetExcelWkbObject(CallingApplication As Excel.Application, ByRef strWorkbookFilePath As String, Optional bPromptToUseFile As Boolean = False, Optional bSaveOnTerminate As Boolean = False, Optional bApplyLeaveFileOpenOverride As Boolean = False, Optional strPromptDialogTitle As String = "File Path") As Object 'Excel.Workbook
'Public Function GetExcelWkbObject(CallingApplication As Excel.Application, ByRef strWorkbookFilePath As String, ByRef strWorksheetName As String, WorkbookReferenceType As eWorkbookReferenceType, Optional bSaveOnTerminate As Boolean = True, Optional bAllowValidWorkbookSelection As Boolean = True, Optional bApplyLeaveFileOpenOverride As Boolean = False) As Object 'Excel.Workbook
'If no path specified or invalid path then Nothing is returned.
'bApplyLeaveFileOpenOverride - allows for file to be left open if Excel was open initially (visible).
    Dim bEvents As Boolean

    On Error GoTo errsub
    Dim lPreviousCursor As Long
    Dim oCommonDialog As clsCommonDlg
    
    SaveOnTerminate = bSaveOnTerminate
    m_LeaveFileOpenOverride = bApplyLeaveFileOpenOverride
    
    Set HostApplication = CallingApplication
    Set m_colWorkbooks = New Collection
    Set oCommonDialog = New clsCommonDlg
        
    ' Check is Microsoft Excel is running.
    m_ExcelWasAlreadyOpen = IsProcessRunning("EXCEL.EXE")
    strWorkbookFilePath = Trim(strWorkbookFilePath)
    
    'Get Excel Application Object
    If m_ExcelWasAlreadyOpen = True Then 'Excel already running use this instance.
        If HostApplication = "Microsoft Excel" Then 'Running this code from within Excel - In process
            Set ApplicationInstance = CallingApplication
            m_ExcelInstanceVisible = ApplicationInstance.Visible
        Else 'Running this code from another process - Out of Process (will be slower).
            Set ApplicationInstance = GetObject(, "EXCEL.APPLICATION") 'Get existing object in local instance.
            Call RegisterInROT(ApplicationInstance.hWnd)

            'Hide to avoid user interaction
            m_ExcelInstanceVisible = ApplicationInstance.Visible
            ApplicationInstance.Visible = msoFalse
        End If
        'Keep track of original worksheets
        Call SaveWorkbooksCollectionByName(ApplicationInstance)
    Else 'Excel not running
        Set ApplicationInstance = CreateObject("EXCEL.APPLICATION") 'Create new object
        Call RegisterInROT(ApplicationInstance.hWnd)
        m_ExcelInstanceVisible = ApplicationInstance.Visible
    'There is a third case that would be to load a secondary existing object that contains the loaded file.  Currently not supported because it is slow.  Instead prompting user close file.
    'Could detect by comparing if m_ExcelWasAlreadyOpen and CallingApplication.hwnd <> m_HostApplication.hwnd then open in secondary instance.
    End If
    
    If oCommonDialog.GetValidFile(strWorkbookFilePath, eFilterFileType.Excel2007, strPromptDialogTitle, bPromptToUseFile) <> vbNullString Then
        'http://support.microsoft.com/kb/825939 - AutomationSecurity
        'http://www.excelforum.com/excel-programming/384353-using-vba-to-disable-macros-when-opening-files.html
        'http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel._application.automationsecurity(office.11).aspx
        m_Automation = ApplicationInstance.AutomationSecurity
'        Dim bAskToUpdateLinks As Boolean: bAskToUpdateLinks = m_HostApplication.AskToUpdateLinks
'        m_HostApplication.AskToUpdateLinks = False 'Suppress "Update Links" dialog (doesn't seem to work)
        bEvents = ApplicationInstance.EnableEvents
        
        If bSuppressVBA = True Then
            'Have seen the opening of the workbook kill message handlers, fix is to diable macros (select True) then open.
            'Select true to
            ApplicationInstance.AutomationSecurity = msoAutomationSecurityForceDisable 'Attempt to suppress VBA macro code (actually depends on Enable.Events setting)
            ApplicationInstance.EnableEvents = False 'Suppress VBA macro code that fires based on events.
        Else
            ApplicationInstance.AutomationSecurity = msoAutomationSecurityLow 'Attempt to suppress VBA macro code (actually depends on Enable.Events setting)
            ApplicationInstance.EnableEvents = True 'Enable VBA macro code that fires based on events.
        End If
        
        ApplicationInstance.DisplayAlerts = False 'Suppress "Update Links" dialog
        
        'Will return new workbook, or existing if already open.
        Set WorkbookInstance = ApplicationInstance.Workbooks.Open(strWorkbookFilePath)   'load file
        Set GetExcelWkbObject = WorkbookInstance
        
        ApplicationInstance.DisplayAlerts = True
    ElseIf bCreateNewWorkbookInstance = True Then
        Set WorkbookInstance = ApplicationInstance.Workbooks.Add 'Create new Excel Workbook
        Set GetExcelWkbObject = WorkbookInstance
    Else
        Set GetExcelWkbObject = Nothing 'return nothing for workbook. Could also return ActiveWorkbook
        m_SaveOnTerminate = False
        Call MsgBox("Error connecting to file:" & vbCrLf & strWorkbookFilePath, vbExclamation, "Error")
    End If
Exit Function

errsub:
    Set oCommonDialog = Nothing
    Call UserForm_Terminate
    
'    If Not WorkbookInstance Is Nothing Then
'        WorkbookInstance.Close
'    End If
'
'    Set GetExcelWkbObject = Nothing
End Function

Public Function CloseExcelWkbObject()
    Dim bEvents As Boolean
    
    If Not ApplicationInstance Is Nothing Then
        bEvents = ApplicationInstance.EnableEvents
        'GBH - should we still execute close code even if workbook is not valid?
        If IsWorkbookValid = True Then
    '        If m_HostApplication.Visible = True Then
                If m_SaveOnTerminate = True And WorkbookInstance.ReadOnly = False Then
            '        m_HostApplication.Windows(WorkbookInstance.Name).Visible = True    'This can get set to invisible so that workbook is hidden the next time it is opened.
                    ApplicationInstance.DisplayAlerts = False 'Suppress file overwrite message.
                    Call WorkbookInstance.Save
                    ApplicationInstance.DisplayAlerts = True
                Else
                    WorkbookInstance.Saved = True 'So that error message does not show if file not saved.
                End If
            
                ApplicationInstance.EnableEvents = True
            
                'If sheet was previously open leave it open, otherwise close it, unless overrride.  Since we have the application object before the sheet is accessed/loaded, can use this method.
                If m_LeaveFileOpenOverride = True Then 'Leave it open
                    ApplicationInstance.Visible = True 'Warning: when workbook is closed, application closes.
                    ApplicationInstance.UserControl = True 'http://msdn2.microsoft.com/en-us/library/aa814561(VS.85).aspx 'Makes file behave as if opened by user.
                ElseIf WasWorkbookOpen(WorkbookInstance) = False Then 'If we opened workbook close it.
                    ApplicationInstance.DisplayAlerts = False  'Hide messages such as: large amount of data on clipboard warning
                    'If Excel was already open this can leave an instance of workbook running which gets released when object is released.
                    'Events on so that Custom Command Bar of ThisWorkbook reloads via events.
                    WorkbookInstance.Close
                    Set WorkbookInstance = Nothing
                    
                    If m_ExcelWasAlreadyOpen = True And ApplicationInstance.EnableEvents = False Then 'Put back to what it was if needed.
                        ApplicationInstance.EnableEvents = bEvents
                    End If
                    
                    ApplicationInstance.DisplayAlerts = True
                End If
    '        End If
    
            'Put back original settings
            ApplicationInstance.AutomationSecurity = m_Automation
        Else
            If Not WorkbookInstance Is Nothing Then
                Set WorkbookInstance = Nothing
            End If
        End If
        
        If m_ExcelWasAlreadyOpen = False Then   'Close the instance of the application that was opened.
            Call ApplicationInstance.Quit
            'Make sure it is dead as above call doesn't always work.
            Call KillProcessByHwndApplicationSimple(ApplicationInstance.hWnd)
        Else
            'Put back original settings
            ApplicationInstance.Visible = m_ExcelInstanceVisible
        End If
        
        Set ApplicationInstance = Nothing
    Else
        Set ApplicationInstance = Nothing
    End If
End Function

Private Function RegisterInROT(hWnd As Long)
    'Use the SendMessage API function to enter Application into Running Object Table.
    'If multiple instances are running, only the first launched is entered, so enter in ROT.
    Call SendMessage(hWnd, WM_USER + 18, 0, 0)
End Function

Private Function IsWorkbookValid() As Boolean
'If visible and user interacts with the WorkbookInstance workbook reference can be invalid.
    Dim str As String
    On Error GoTo errsub
    str = WorkbookInstance.name 'Test to see if object is valid.
    IsWorkbookValid = True
errsub:
End Function

Private Function WasWorkbookOpen(wkbQuery As Object) As Boolean
'If workbook was originally in application then it was not opened and the existing version is being used.
    Dim wkb As Variant 'String
    
    For Each wkb In m_colWorkbooks
        If wkb = wkbQuery.name Then 'Look for worksheet by name
            WasWorkbookOpen = True 'Falls through when error occurs (value not found in collection)
            Exit For
        End If
    Next
End Function

Private Function SaveWorkbooksCollectionByName(SourceApplication As Object) 'Application)
    Dim wkb As Object 'Excel.Workbook
    
    For Each wkb In SourceApplication.Workbooks
        m_colWorkbooks.Add wkb.name, wkb.name
    Next
End Function
        
Private Function IsProcessRunning(ByVal strApp As String) As Boolean
    IsProcessRunning = KillProcessByNameWMI(strApp, False)
End Function

Public Function KillProcessByNameWMI(procName As String, bKillProcess As Boolean) As Boolean
'Kills all occurances of an application in the process list if bKillProcess is True
'If bKillProcess is false, returns if process is running, without killing it.
'Uses: Windows Management Instrumentation - Microsoft WMI Scripting v1.2 Scripting
    On Error GoTo 0
    Dim objProcList As Object 'SWbernObjectSet
    Dim objWMI As Object 'SWbernServicesEx
    Dim objProc As Object 'SWbernObjectEx
    
    'create WMI object instance
    Set objWMI = GetObject("winmgmts:")
'    Debug.Print "Cleaning up " & procName
    If Not IsNull(objWMI) Then
        'create object collection of Win32 processes
        Set objProcList = objWMI.ExecQuery("SELECT * FROM Win32_Process where Caption='" & procName & "'")
'        Set objProcList = objWMI.InstancesOf("win32_process")
        For Each objProc In objProcList 'iterate through enumerated collection
            If UCase(objProc.name) = UCase(procName) Then  'Double check
                KillProcessByNameWMI = True
                If bKillProcess = True Then
                    objProc.Terminate (0)
                    Debug.Print procName & " was terminated"
                End If
            End If
        Next
    End If
    Set objProcList = Nothing
    Set objWMI = Nothing
End Function

Private Function KillProcessByHwndApplicationSimple(hWndApplication As Long) As Boolean
'Used to kill a particular instance of an application by PID (processor ID).
'Modified from:
'http://www.microsoft.com/technet/scriptcenter/resources/qanda/sept04/hey0927.mspx

    On Error GoTo 0
    Dim objProcList As Object 'SWbernObjectSet
    Dim objWMI As Object 'SWbernServicesEx
    Dim objProc As Object 'SWbernObjectEx
    Dim lRet As Long
    Dim lProcessID As Long
    
    lRet = GetWindowThreadProcessId(hWndApplication, lProcessID)
        
    'create WMI object instance
    Set objWMI = GetObject("winmgmts:")
    
    If Not IsNull(objWMI) And lRet > 0 Then
        Set objProcList = objWMI.ExecQuery("Select * from Win32_Process Where ProcessID = " & lProcessID & "")
        
'        Debug.Print "Ending process: Excel.exe with Process ID: " & lProcessID
        
        For Each objProc In objProcList
            objProc.Terminate (0)
        Next
    End If
    
    Set objProc = Nothing
    Set objProcList = Nothing
    Set objWMI = Nothing
End Function

Public Function GetNewWorksheet(strWorksheetName As String, bDeleteExtraWorksheets As Boolean) As Object 'Excel.worksheet
'Inserts new worksheet in first position, optionaly removing extras.
    Dim bPrevious As Boolean
    
    bPrevious = ApplicationInstance.DisplayAlerts
    
    ApplicationInstance.DisplayAlerts = False
    Call WorkbookInstance.Worksheets.Add(WorkbookInstance.Worksheets(1)) 'Add new worksheet to first position.
    If bDeleteExtraWorksheets = True Then
        Do While WorkbookInstance.Worksheets.Count > 1
            WorkbookInstance.Worksheets(2).Delete
        Loop
    End If
    
    WorkbookInstance.Worksheets(1).name = Trim(VBA.Left(strWorksheetName, 30)) 'Worksheet name limit
    Set GetNewWorksheet = WorkbookInstance.Worksheets(1)
    
errsub:
    ApplicationInstance.DisplayAlerts = bPrevious
End Function
    
Public Function GetValidWorksheet(strWorksheetName As String, WorkbookReferenceType As eWorkbookReferenceType) As Object 'Excel.Worksheet
    If WorkbookReferenceType = eWorkbookReferenceType.WorksheetName Then
        Set GetValidWorksheet = GetValidWorksheetByName(strWorksheetName, True)
    Else
        Set GetValidWorksheet = GetValidWorksheetByCodeName(strWorksheetName, True)
    End If
End Function

Private Function GetValidWorksheetByName(strWorksheetName As String, bAllowValidWorkbookSelection As Boolean) As Object 'Excel.Worksheet
'Returns worksheet object. Set to nothing if not found.
    Dim wkbSource As Object 'Excel.Workbook
    Dim wksFound As Object 'Excel.Worksheet
    Dim wks As Object 'Excel.Worksheet
    Dim bFound As Boolean
    Dim bScreenUpdating As Boolean
    
    Set wkbSource = WorkbookInstance
        
    'Populate for possible use here and other use.  Must be at least one sheet.
    Me.cmbWorksheets.Clear
    If Not wkbSource Is Nothing Then
        
        For Each wks In wkbSource.Worksheets
            Me.cmbWorksheets.AddItem (wks.name)
            If wks.name = strWorksheetName Then
                bFound = True
                Set wksFound = wks
            End If
        Next
        Me.cmbWorksheets.ListIndex = 0 'set to first item in list.
    
        'Get worksheet name.  If not found optionally allow for selection.
        If bFound = False Then
            If bAllowValidWorkbookSelection = True Then
                If HostIsExcel Then
                    bScreenUpdating = HostApplication.ScreenUpdating 'Allow painting while dialog open.
                    HostApplication.ScreenUpdating = True
                End If
                
                Me.Show vbModal 'Unload of form canceled with QueryClose so that we don't lose object references.
                Set wksFound = GetWorksheetByName(wkbSource, Me.cmbWorksheets.Text)
                
                If HostIsExcel Then
                    HostApplication.ScreenUpdating = bScreenUpdating
                End If
'                Set wks = wkbSource.Worksheets(Me.cmbWorksheets.Text)
    '        Else
                'Could set to first valid sheet name (there must be at least one in Excel workbook)
    '            Set wks = Nothing 'wkbSource.Worksheets(Me.cmbWorksheets.List(0))
            End If
        End If
        
        If Not wksFound Is Nothing Then
            m_ExcelWorksheetName = wksFound.name
            strWorksheetName = wksFound.name
            Set GetValidWorksheetByName = wksFound
        End If
    End If
End Function

Private Function GetValidWorksheetByCodeName(strWkbCodeName As String, bAllowValidWorkbookSelection As Boolean) As Object 'Excel.Worksheet
'Public Function GetValidWorksheetByCodeName(wkbSource As Workbook, ByRef strWkbCodeName As String, bAllowValidWorkbookSelection As Boolean) As Excel.Worksheet
'Returns worksheet object. Set to nothing if not found.
    Dim wkbSource As Object 'Workbook
    Dim wksFound As Object 'Excel.Worksheet
    Dim wks As Object 'Excel.Worksheet
    Dim bFound As Boolean
    Dim bScreenUpdating As Boolean
    
    Set wkbSource = WorkbookInstance
    
    'Populate for possible use here and other use.  Must be at least one sheet.
    Me.cmbWorksheets.Clear
    
    If Not wkbSource Is Nothing Then
        For Each wks In wkbSource.Worksheets
            Me.cmbWorksheets.AddItem (wks.CodeName)
            If wks.CodeName = strWkbCodeName Then
                bFound = True
                Set wksFound = wks
            End If
        Next
        Me.cmbWorksheets.ListIndex = 0 'set to first item in list.
    
        'Get worksheet name.  If not found optionally allow for selection.
        If bFound = False Then
            If bAllowValidWorkbookSelection = True Then
                If HostIsExcel Then
                    bScreenUpdating = HostApplication.ScreenUpdating 'Allow painting while dialog open.
                    HostApplication.ScreenUpdating = True
                End If
                
                Me.Show vbModal 'Unload of form canceled with QueryClose so that we don't lose object references.
                Set wksFound = GetWorksheetByCodeNameString(Me.cmbWorksheets.Text)
                
                If HostIsExcel Then
                    HostApplication.ScreenUpdating = bScreenUpdating
                End If
    '            Set wks = wkbSource.Worksheets(Me.cmbWorksheets.Text)
    '        Else
                'Could set to first valid sheet name (there must be at least one in Excel workbook)
    '            Set wks = Nothing 'wkbSource.Worksheets(Me.cmbWorksheets.List(0))
            End If
        End If
        
        If Not wksFound Is Nothing Then
            m_ExcelWorksheetName = wksFound.name
            strWkbCodeName = wksFound.name
            Set GetValidWorksheetByCodeName = wksFound
        End If
    End If
End Function

Public Function GetWorkbookProperty(wkbActive As Object, strPropertyName As String, Optional bIsCustomProperty As Boolean = True) As Variant
'Public Function GetWorkbookProperty(wkbActive As Excel.Workbook, strPropertyName As String, Optional bIsCustomProperty As Boolean = True) As Variant
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
'Public Sub SetWorkbookProperty(wkbActive As Excel.Workbook, strPropertyName As String, vPropertyValue As Variant, Optional bIsCustomProperty As Boolean = True)
'Public Sub SetProperty(WorkbookName As String, PropName As String, PValue As Variant, PropCustom As Boolean)
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

Private Function GetWorksheetByName(wkbSource As Object, strSheetName As String) As Object
'Private Function GetWorksheetByName(wkbSource As Excel.Workbook, strSheetName As String) As Excel.Worksheet
'Returns worksheet object. Set to nothing if not found.
    On Error Resume Next
    With wkbSource
    '    Worksheets (strSheetName)
        Set GetWorksheetByName = .Worksheets(strSheetName)
        If Err.Number <> 0 Then
            Set GetWorksheetByName = Nothing
        End If
    End With
    Err.Clear
End Function

Public Function GetWorksheetByCodeNameString(strSheetCodeName As String) As Object 'Excel.Worksheet
'Public Function GetWorksheetByCodeNameString(strSheetCodeName As String) As Excel.Worksheet
'Used to get worksheet object by codename string
    On Error GoTo errsub

    Dim sht As Object 'Excel.Worksheet
    For Each sht In WorkbookInstance.Worksheets
        If sht.CodeName = strSheetCodeName Then
            Set GetWorksheetByCodeNameString = sht
            Exit Function
        End If
    Next
errsub:
    Err.Clear
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''UserForm Code
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserForm_Initialize()   'Private Sub Class_Initialize()
    Dim oPW As clsPositionWindow
    Set oPW = New clsPositionWindow
    
    Me.Hide
    'Set default position
    Call oPW.ForceWindowIntoWorkArea(Me)

    Set oPW = Nothing
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then   ' user clicked the X button
        ' Cancel unloading the form object
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub UserForm_Terminate()    'Private Sub Class_Terminate()
    Call CloseExcelWkbObject
End Sub

Private Function HostIsExcel() As Boolean
    If HostApplication.name = "Microsoft Excel" Then
        HostIsExcel = True
    End If
End Function

'Example From Help:  Might help to improve code above.  Problem is that when using GetObject, the object has to be set to visible before closing (or it is permanently invisible) which is not desirable.
'example
'This example uses the GetObject function to get a reference to a specific Microsoft Excel worksheet (
'
'MyXL
'
'). It uses the worksheet's Application property to make Microsoft Excel visible, to close it, and so on.
'Using two API calls, the DetectExcel Sub procedure looks for Microsoft Excel, and if it is running, enters it in the Running Object Table.
'The first call to GetObject causes an error if Microsoft Excel isn't already running.
'In the example, the error causes the ExcelWasNotRunning flag to be set to True.
'The second call to GetObject specifies a file to open. If Microsoft Excel isn't already running, the second call starts it and returns a reference to the worksheet represented by the specified file, mytest.xls.
'The file must exist in the specified location; otherwise, the Visual Basic error Automation error is generated.
'Next the example code makes both Microsoft Excel and the window containing the specified worksheet visible.
'Finally, if there was no previous version of Microsoft Excel running, the code uses the Application object's Quit method to close Microsoft Excel. If the application was already run
'ing, no attempt is made to close it. The reference itself is released by setting it to Nothing.
'
'' Declare necessary API routines:
'Declare Function FindWindow Lib "user32" Alias _
'"FindWindowA" (ByVal lpClassName As String, _
'                    ByVal lpWindowName As Long) As Long
'
'Declare Function SendMessage Lib "user32" Alias _
'"SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
'                    ByVal wParam As Long, _
'                    ByVal lParam As Long) As Long
'
'Sub GetExcel()
'    Dim MyXL As Object    ' Variable to hold reference
'                                ' to Microsoft Excel.
'    Dim ExcelWasNotRunning As Boolean    ' Flag for final release.
'
'' Test to see if there is a copy of Microsoft Excel already running.
'    On Error Resume Next    ' Defer error trapping.
'' Getobject function called without the first argument returns a
'' reference to an instance of the application. If the application isn't
'' running, an error occurs.
'    Set MyXL = GetObject(, "Excel.Application")
'    If Err.Number <> 0 Then ExcelWasNotRunning = True
'    Err.Clear    ' Clear Err object in case error occurred.
'
'' Check for Microsoft Excel. If Microsoft Excel is running,
'' enter it into the Running Object table.
'    DetectExcel
'
'' Set the object variable to reference the file you want to see.
'    Set MyXL = GetObject("c:\vb4\MYTEST.XLS")
'
'' Show Microsoft Excel through its Application property. Then
'' show the actual window containing the file using the Windows
'' collection of the MyXL object reference.
'    MyXL.Application.Visible = True
'    MyXL.Parent.Windows(1).Visible = True
'     Do manipulations of your  file here.
'    ' ...
'' If this copy of Microsoft Excel was not running when you
'' started, close it using the Application property's Quit method.
'' Note that when you try to quit Microsoft Excel, the
'' title bar blinks and a message is displayed asking if you
'' want to save any loaded files.
'    If ExcelWasNotRunning = True Then
'        MyXL.Application.Quit
'    End If
'
'    Set MyXL = Nothing    ' Release reference to the
'                                ' application and spreadsheet.
'End Sub
'
'Sub DetectExcel()
'' Procedure dectects a running Excel and registers it.
'    Const WM_USER = 1024
'    Dim hWnd As Long
'' If Excel is running this API call returns its handle.
'    hWnd = FindWindow("XLMAIN", 0)
'    If hWnd = 0 Then    ' 0 means Excel not running.
'        Exit Sub
'    Else
'    ' Excel is running so use the SendMessage API
'    ' function to enter it in the Running Object Table.
'        SendMessage hWnd, WM_USER + 18, 0, 0
'    End If
'End Sub
