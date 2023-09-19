VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clsProject 
   Caption         =   "Project Selection"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6675
   OleObjectBlob   =   "clsProject.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clsProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'File:   clsProject
'Author:    Greg Harward
'Contact:   gharward@gmail.com
'Date:      1/17/16

'Notes:
'To convert between field constants and names:
'FieldNameToFieldConstant()
'FieldConstantToFieldName()

'Private Sub GetTaskNames()
'    Dim o As Object
'    For Each o In ActiveProject.TaskTables("Entry").TableFields
'        MsgBox FieldConstantToFieldName(o.Field)
'    Next
'End Sub

'Summary:
'Created to handle the acquire and release of the Project object which is a multi-use COM component (only a single instance can be open at one time).
'Make sure to check the return object is Not Nothing when using.
'Project only allows for a single instance to be open at a time.
'If file path is valid then file is then opened in application instance acquired.
'If file is already opened and on remote application instance, prompts to close the file first.
'This could be replaced with ability to access the remote instance, however using this method of access is very slow so prompting instead.
'If adding this ability would then also need to check on workbook close if app should also be closed.
'Code loses object handle (m_prjFile) if workbook is made visible and user interacts with the workbook.  Object is not nothing, but is also not filled.

'Project VBA References:
'https://msdn.microsoft.com/en-us/library/office/ee861523.aspx
'http://ptgmedia.pearsoncmg.com/images/0789727013/downloads/7013Web2.PDF

'Addresses Microsoft bug described here:
'PRB: Releasing Object Variable Does Not Close Microsoft Excel
'http://support.microsoft.com/kb/132535/EN-US
'http://msdn.microsoft.com/archive/default.asp?url=/archive/en-us/dnaraccessdev/html/ODC_MicrosoftAccessOLEAutomation.asp
'http://msdn2.microsoft.com/en-us/library/e9waz863(VS.71).aspx

''Example Use:
''Private Sub BuildPortfolioBuilderTemplate()
''    Dim wkb As Excel.Workbook
''    Dim strExcelSourceFile As String
''    Dim ExcelWrapper As clsExcel
''    Set ExcelWrapper = New clsExcel
''
''    Set wkb = ExcelWrapper.GetExcelWkbObject(Me.Application, strExcelSourceFile, True)
''    If Not wkb Is Nothing Then
''        '<Code Here>
''        Call ExcelWrapper.CloseExcelWkbObject
''    End If
''End Sub

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

'Did not include code for restricting screen updating or changing curors such as:
'    Application.ScreenUpdating = False
'    Application.Cursor = xlWait

Private Const pjDoNotSave = 0
Private Const pjSave = 1

Private Const WM_USER = 1024
Private Const prjObjMain = "JWinproj-WhimperMainClass"                       'Project Application

Public Enum eWorkbookReferenceType
    WorksheetName = 0
    WorksheetCodeName = 1
End Enum

'Use SendMessageTimeout() instead of SendMessage() so that application doesn't hang waiting for response if destination window is unresponsive.
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As String, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwprocessid As Long) As Long
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long

Private m_ProjectWasAlreadyOpen As Boolean
Private m_SaveOnTerminate As Boolean 'Application
Private m_SaveOnClose As Boolean 'Project File
Private m_LeaveFileOpenOverride As Boolean

Private m_HostApplication As Object 'Application calling Project instance.
Private m_ApplicationInstance As Object 'Project.Application
Private m_prjFile As Object 'Project File

'Private m_ExcelWorksheetName As String
Private m_colProjects As Collection
Private m_Automation As Office.MsoAutomationSecurity
Private m_ProjectInstanceVisible As Boolean

'Private Const xlWait = 2
'Private Const xlDefault = &HFFFFEFD1    '4143

Public Property Get ApplicationInstance() As Object
    Set ApplicationInstance = m_ApplicationInstance
End Property

Private Property Set ApplicationInstance(ByRef Instance As Object)
    Set m_ApplicationInstance = Instance
End Property

Public Property Get ProjectFileInstance() As Object
    Set ProjectFileInstance = m_prjFile
End Property

Private Property Set ProjectFileInstance(ByRef Workbook As Object)
    Set m_prjFile = Workbook
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

Public Function GetProjectObject(CallingApplication As Application, Optional bSaveOnTerminate As Boolean = False, Optional bApplyLeaveFileOpenOverride As Boolean = False, Optional bSuppressVBA = True) As Boolean
'Return Project Application object and sets ApplicationInstance Property
'If no path specified or invalid path then Nothing is returned.
'bApplyLeaveFileOpenOverride - allows for file to be left open if CallingApplication was open initially (visible).
    On Error GoTo errsub
    
    Dim lPreviousCursor As Long
     
    SaveOnTerminate = bSaveOnTerminate
    m_LeaveFileOpenOverride = bApplyLeaveFileOpenOverride
    
    Set HostApplication = CallingApplication
    Set m_colProjects = New Collection
        
    ' Check if already running.
    m_ProjectWasAlreadyOpen = IsProcessRunning("WINPROJ.EXE")
   
    'Get Application Object
    If m_ProjectWasAlreadyOpen = True Then 'Already running use this instance.
        If HostApplication = "Microsoft Project" Then 'Running this code in this process
            Set ApplicationInstance = CallingApplication
            m_ProjectInstanceVisible = ApplicationInstance.Visible
        Else 'Running this code from another process - Out of Process (will be slower).
            Set ApplicationInstance = GetObject(, "MSPROJECT.APPLICATION") 'Get existing object in local instance.
            Call RegisterInROT(GetApplicationHwnd())

            'Hide to avoid user interaction
            m_ProjectInstanceVisible = ApplicationInstance.Visible
'            ApplicationInstance.Visible = msoFalse
        End If
        'Keep track of original projects
        Call SaveProjectsCollectionByName(ApplicationInstance)
    Else 'Not running
        Set ApplicationInstance = CreateObject("MSPROJECT.APPLICATION") 'Create new object
        Call RegisterInROT(GetApplicationHwnd())
        m_ProjectInstanceVisible = ApplicationInstance.Visible
    'There is a third case that would be to load a secondary existing object that contains the loaded file.  Currently not supported because it is slow.  Instead prompting user close file.
    'Could detect by comparing if m_ProjectWasAlreadyOpen and CallingApplication.hwnd <> m_HostApplication.hwnd then open in secondary instance.
    End If

    If bSuppressVBA = True Then
        'Have seen the opening of the projects kill message handlers, fix is to diable macros (select True) then open.
        ApplicationInstance.AutomationSecurity = msoAutomationSecurityForceDisable 'Attempt to suppress VBA macro code (actually depends on Enable.Events setting)
    Else
        ApplicationInstance.AutomationSecurity = msoAutomationSecurityLow 'Attempt to suppress VBA macro code (actually depends on Enable.Events setting)
    End If
    
    'http://support.microsoft.com/kb/825939 - AutomationSecurity
     'http://www.excelforum.com/excel-programming/384353-using-vba-to-disable-macros-when-opening-files.html
     'http://msdn.microsoft.com/en-us/library/microsoft.office.interop.excel._application.automationsecurity(office.11).aspx
    m_Automation = ApplicationInstance.AutomationSecurity
'     Dim bAskToUpdateLinks As Boolean: bAskToUpdateLinks = m_HostApplication.AskToUpdateLinks
'     m_HostApplication.AskToUpdateLinks = False 'Suppress "Update Links" dialog (doesn't seem to work)

    GetProjectObject = True 'ApplicationInstance
    
'    If strFilePath <> vbNullString Then
'        Set GetProjectObject = OpenProjectFileObject(strFilePath, bPromptToUseFile)
'    End If
    
    Exit Function
errsub:
    Call UserForm_Terminate
End Function

Public Function OpenProjectFileObject(ByRef strFilePath As String, Optional bPromptToUseFile As Boolean = False, Optional strPromptDialogTitle As String = "File Path") As Object
    On Error GoTo errsub
    
    Dim bDisplayAlerts As Boolean
    Dim oProjectFile As Object
    Dim oCommonDialog As clsCommonDlg
    Set oCommonDialog = New clsCommonDlg
    
    If oCommonDialog.GetValidFile(strFilePath, eFilterFileType.Project, strPromptDialogTitle, bPromptToUseFile) <> vbNullString Then
        bDisplayAlerts = ApplicationInstance.DisplayAlerts
        ApplicationInstance.DisplayAlerts = False 'Suppress "Update Links" dialog
        
        If ApplicationInstance.FileOpenEx(strFilePath, True) = True Then
            Set ProjectFileInstance = ApplicationInstance.ActiveProject
            Set OpenProjectFileObject = ProjectFileInstance
'            ApplicationInstance.Visible = True
            If ApplicationInstance.Visible = True Then
                Call ToggleWindowVisibility(ApplicationInstance, ProjectFileInstance.FullName, True)
            End If
        End If
    Else
'        m_SaveOnTerminate = False
        m_SaveOnClose = False
        If bPromptToUseFile = True Then
            Call MsgBox("Error connecting to file:" & vbCrLf & strFilePath, vbExclamation, "Error")
        End If
    End If
    
errsub:
    ApplicationInstance.DisplayAlerts = bDisplayAlerts
    Set oCommonDialog = Nothing
End Function

Public Function CloseProjectObject()
    Dim bEvents As Boolean
    
    If Not ApplicationInstance Is Nothing Then
        If IsProjectValid = True Then
            If m_SaveOnTerminate = True Then 'And WorkbookInstance.ReadOnly = False Then
        '        m_HostApplication.Windows(WorkbookInstance.Name).Visible = True    'This can get set to invisible so that workbook is hidden the next time it is opened.
                ApplicationInstance.DisplayAlerts = False 'Suppress file overwrite message.
                Call m_prjFile.Save
                ApplicationInstance.DisplayAlerts = True
            End If
        End If
        
        If m_ProjectWasAlreadyOpen = False Then   'Close the instance of the application that was opened.
            Call ApplicationInstance.Quit
            'Make sure it is dead as above call doesn't always work.
            Call KillProcessByHwndApplicationSimple(GetApplicationHwnd())
        Else
            'Restore original settings
            ApplicationInstance.Visible = m_ProjectInstanceVisible
            ApplicationInstance.AutomationSecurity = m_Automation
        End If
        
        Set ApplicationInstance = Nothing
    Else
        Set ApplicationInstance = Nothing
    End If
End Function

Private Function ToggleWindowVisibility(oApplication As Object, strCaption As String, Optional Hidden = False) As Boolean
    On Error GoTo errsub
    
    If Hidden = True Then
        Call oApplication.WindowActivate(strCaption) 'Also unhides window
        Call oApplication.WindowHide(strCaption)
    Else
        Call oApplication.WindowUnhide(strCaption)
        Call oApplication.WindowActivate(strCaption) 'Also unhides window
    End If
    ToggleWindowVisibility = True
errsub:
End Function
                
Private Function GetWindowByCaption(strCaption As String) As Object 'Window
'Workaround to get window object by name. Not working by direct reference.
    Dim oWindow As Object
    For Each oWindow In ApplicationInstance.Windows
        If oWindow.Caption = strCaption Then
            Set GetWindowByCaption = oWindow
            Exit Function
        End If
    Next
End Function

Public Function CloseProjectFile(oProjectFile As Object) As Boolean
            On Error GoTo errsub
            'If project was previously open leave it open, otherwise close it, unless overrride.
            If m_LeaveFileOpenOverride = True Then 'Force leave it open
                ApplicationInstance.Visible = True 'Warning: when workbook is closed, application closes.
'                ApplicationInstance.UserControl = True 'http://msdn2.microsoft.com/en-us/library/aa814561(VS.85).aspx 'Makes file behave as if opened by user.
                Call ToggleWindowVisibility(ApplicationInstance, oProjectFile.FullName)
            ElseIf WasProjectFileOpen(oProjectFile) = False Then 'If we opened workbook close it.
                ApplicationInstance.DisplayAlerts = False  'Hide messages such as: 'large amount of data on clipboard' warning
                'https://msdn.microsoft.com/EN-US/library/office/ff863262.aspx
                Call ToggleWindowVisibility(ApplicationInstance, oProjectFile.FullName)
                Call ApplicationInstance.FileCloseEx(pjDoNotSave)
                ApplicationInstance.DisplayAlerts = True
            Else 'It was open when we started, so leave it open and make sure that it is showing.
                Call ToggleWindowVisibility(ApplicationInstance, oProjectFile.FullName)
            End If
            CloseProjectFile = True
errsub:
End Function

Private Function RegisterInROT(hWnd As Long)
    'Use the SendMessage API function to enter Application into Running Object Table.
    'If multiple instances are running, only the first launched is entered, so enter in ROT.
    Call SendMessage(hWnd, WM_USER + 18, 0, 0)
End Function

Private Function IsProjectValid() As Boolean
'If visible and user interacts with the WorkbookInstance workbook reference can be invalid.
    Dim str As String
    On Error GoTo errsub
    str = ApplicationInstance.name 'Test to see if object is valid.
    IsProjectValid = True
errsub:
End Function

Private Function WasProjectFileOpen(prjQuery As Object) As Boolean
'If workbook was originally in application then it was not opened and the existing version is being used.
    Dim prj As Variant 'String

    For Each prj In m_colProjects
        If prj = prjQuery.name Then 'Look for file by name
            WasProjectFileOpen = True 'Falls through when error occurs (value not found in collection)
            Exit For
        End If
    Next
End Function

Private Function SaveProjectsCollectionByName(SourceApplication As Object) 'Application)
    Dim prj As Object 'Excel.Workbook

    For Each prj In SourceApplication.Projects
        m_colProjects.Add prj.name, prj.name
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
            If VBA.UCase(objProc.name) = VBA.UCase(procName) Then  'Double check
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

Private Function GetApplicationHwnd() As Long
    On Error GoTo errsub
    GetApplicationHwnd = FindWindowEx(vbEmpty, vbEmpty, prjObjMain, ApplicationInstance.Caption)
errsub:
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''UserForm Code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    Call CloseProjectObject
End Sub
