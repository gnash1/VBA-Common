VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} clsDBConnectivity 
   Caption         =   "Database Connection"
   ClientHeight    =   1005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6675
   OleObjectBlob   =   "clsDBConnectivity.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "clsDBConnectivity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'File:   clsDBConnectivity
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Date:        12/25/2018
'Class for management of DB connections within MS Office Applications.

'Summary:
'Code to connect to various datasources using JET OLE DB Drivers to manage data connection/interaction.
'When OLE DB drivers are not available Jet or DSN-less ODBC is used.
'(Note: This only works for ADO 2.0 and up! If you are using an older version of ADO, you will need to upgrade. You can download the latest version of ADO for free at http://www.microsoft.com/data.)
'Jet Drivers are limited to 50(Jet3.5), 99 (Jet4.0), 127 fields (depending on the driver) see http://support.microsoft.com/kb/192716
'JET is inherently limited to 2G for data size.
'Jet ISAM names implemented in GetISAMName() that allow DB to DB data transfer.  Use ISAM name in place of DB name in SQL Query.
'Oracle assumes that the files sqlnet.ora and tnsnames.ora are in the Oracle home directory.  Example: 'C:\Oracle\network\Admin
'Connection to currently open Excel files is permitted.  This includes call to file that contains this code itself.
'Connection to Excel Uses Microsoft.Jet.OLEDB.4.0 driver (which is newer than native ODBC drivers).
'http://support.microsoft.com/kb/316934
'Connection to Excel supports three connection modes which are: Workbook(contains trialing $), Named Range, and Range Address <worksheet name>$<Range:Range>
'Using the Connection.Execute() is more performant then Recordset.Open.
'If early bound, requires reference to Microsoft ActiveX Data Objects 2.8 Library
'C:\Program Files\Common Files\system\ado\msado15.dll
'Warning: Have seen csv truncate floating point numbers to integers, with no apparent solution.

'To Do:
'Replace ability to pick from combo box with default functionality provided by: m_DBConnection.Properties("Prompt") = SQLADOConnectionPromptEnum.adPromptAlways
'Replace GetOpenFilename with common code that can be used by any application. GetOpenFilename only works with Excel.
'Validate  GetValidDataTableName() call to OpenSchema for various data souces.  http://support.microsoft.com/kb/186246
'Fix: m_Application.XlMousePointer.xlWait
'Sort table names before adding to combo.
'If invalid paths passed in (when file based), trapped error occurs.

'Query Creation Steps and conversion to VBA:
'1. In Excel Import Data "From Other Sources" - "From Microsoft Query"
'2. Uncheck box that says: "Use the Query Wizard to create/edit queries" in order to use newer Microsoft Query version 12.
'3. Select Databases - Excel Files (or Queries tab if loading existing)
'4. After query is running select SQL button to copy SQL code from window.
'5. Replace table name if necessary using this format: [tablename$]
'6. Replace ` with [ and ] (used to pass strings with spaces).

'Excel Query Syntax Options:
'To read a sheet:
'"SELECT * FROM [Sheet1$]"
'To refer to a range by its address:
'"SELECT * FROM [Sheet1$A1:D10]"
'To refer to a single-cell range, pretend it's a multi-cell range and specify both the top-left and bottom-right cells:
'"SELECT * FROM [Sheet1$A1:A1]"
'To read a Workbook-level named range:
'"SELECT * FROM MyDataRange"
'To read a Worksheet-level named range
'"SELECT * FROM [Sheet1$MyData]"
'"SELECT * FROM [Sheet1$MyData]" OR "SELECT * FROM ['Sheet Name With Space'$MyData]"

'Testing/Creating OLE DB connection string
'1. Create text file with the extention "*.udl".
'2. double click to open, selecting driver and entering setting.
'3. Select "Test Connection" button.  When successful close dialog.
'4. Open *.udl file as text to see connection string information.

'Online References:
'http://www.connectionstrings.com
'Very good overview of issues related to working with ADO in Excel:
'   http://www.xtremevbtalk.com/showthread.php?t=217783
'http://4guysfromrolla.com/webtech/063099-1.shtml
'http://www.carlprothman.net/Default.aspx?tabid=87#OLEDBProviderForOracleFromOracle
'http://www.carlprothman.net/Default.aspx?tabid=87#OLEDBProviderForExcel
'http://www.carlprothman.net/Default.aspx?tabid=87#OLEDBProviderForMicrosoftJetExcel
'http://www.xtremevbtalk.com/showthread.php?p=969626#post969626
'http://support.microsoft.com/kb/257819
'http://msdn2.microsoft.com/en-us/library/Bb264566.aspx
'http://www.w3schools.com/ADO/met_rs_open.asp
'http://www.beyondtechnology.com/geeks023.shtml 'Example using default ADODB provider.
'http://www.asp101.com/articles/john/connstring/default.asp
'http://support.microsoft.com/default.aspx?scid=kb;en-us;194124 - Explains IMEX
'http://filehelpers.sourceforge.net/ -Free .Net library to import/export fixed length or delimited records
'http://support.microsoft.com/kb/246335 - How to transfer data from an ADO Recordset to Excel with automation
'http://www.w3schools.com/ado/met_rs_getrows.asp#getrowsoptionenum
'http://social.msdn.microsoft.com/Forums/en-US/sqldataaccess/thread/8514b4bb-945a-423b-98fe-a4ec4d7366ea - Info on OPENROWSET with Microsoft.ACE.OLEDB.12.0
'http://www.integralwebsolutions.co.za/Blog/EntryId/283/Importing-and-using-Excel-data-into-MS-SQL-database.aspx - importing Excel data into MSSQL
'http://support.microsoft.com/kb/316934
'Contains workaround for issue described here: http://support.microsoft.com/kb/319998 - Describes memory leak that occures when performing query on initiating Excel workbook using Jet provider.
'http://support.microsoft.com/kb/295646/ - Described how to INSERT and Append using OLEDB
'http://msdn.microsoft.com/en-us/library/ms974559 - query text files using ADO including schema.ini

'Enable ADO Connection Pooling in ADO application.
'http://support.microsoft.com/kb/237844

'Office 2007 ACE.OLEDB Drivers
'http://www.microsoft.com/download/en/details.aspx?displaylang=en&id=23734

'Info on User Instance, Named Pipes, SQL 2005 Express, and MDAC
'http://technet.microsoft.com/en-us/library/ms345154(SQL.90).aspx
'http://msdn.microsoft.com/en-US/library/aa337276(v=SQL.90).aspx
'http://msdn.microsoft.com/en-us/library/ms254504.aspx

'Diagram of ADO, ODBC, OLEDB,.Net architecture.
'http://upload.wikimedia.org/wikipedia/commons/thumb/8/8d/MDAC_Architecture.svg/490px-MDAC_Architecture.svg.png

'C ADO Class
'http://www.codeproject.com/KB/database/caaadoclass1.aspx#OpenXML

'Oracle
'Oracle 11g ODAC and Oracle Developer Tools for Visual Studio (ODAC - Oracle Data Access Components)
'http://www.oracle.com/technology/software/tech/dotnet/utilsoft.html

'Revisions:
'Date     Initials    Description of changes


'Syntax Notes:
'Jet drivers do not support the Select Case Statement - use IIF(), Choose(), or SWITCH() instead.
'VBA functions are supported
'Some report that Excel 2003 need the exta OLEDB; section in the beginning of the connection string such as:
'OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\MyExcel.xls;Extended Properties="Excel 8.0;HDR=Yes;IMEX=1";
'Important note!
'The quota " in the string needs to be escaped using your language specific escape syntax.
'c#, c++   \"
'VB6, VBScript   ""
'xml (web.config etc)   &quot;
'or maybe use a single quota '.
'"HDR=Yes;" indicates that the first row contains columnnames, not data. "HDR=No;" indicates the opposite.
'"IMEX=1;" tells the driver to always read "intermixed" (numbers, dates, strings etc) data columns as text. Note that this option might affect excel sheet write access negative.
'IMEX settings: 0 - export mode, 1 - import mode, 2 - linked mode (full update capabilities)
'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel] REG_DWORD "TypeGuessRows". Specifies range of rows to scan to guess data type (0-16).  Setting to 0 to scan all rows, does nothing.
'Summary of data type guessing problem.  Source is last entry at: http://www.xtremevbtalk.com/showthread.php?t=217783
'Use TypeGuessRows to get Jet to detect whether a 'mixed types ' situation exists or use it to 'trick' Jet into detecting a certaint data type as being the majority type. In the event of a
'mixed types' situation being detected, use ImportMixedTypes to tell Jet to either use the majority type or coerce all values as 'Text' (max 255 characters).
'If the Excel workbook is protected by a password, you cannot open it for data access, even by supplying the correct password with your connection string. If you try, you receive the following error message: "Could not decrypt file."
'When performing server side date calculation query, format as #3/13/08# such as: [CDS_DATE] <= #" & Format(Now(), "m/d/yyyy;@") & "#"
'When fomating dates client side and sending to query format as: "[`" & Format(Now(), "m/d/yyyy") & "`]"
'$ symbol used for sheet names
' "[" & "]" & " ` " (bellow ~) used for strings.  Square brackets are prefered.  Use around portions of query that contain spaces.
'"#" used in place of "." when period is part of a field name.
'char(96) is used surrounding dates to query special field syntax.  Character is: ` Example: When performing query against date field such as '5/31/2010'
'When querying date field names from Excel query as numbers such as: "SELECT [40299] AS [5/1/2010] FROM [WF Demand$]"  It was also seen to be queried as date string.  Match header values types after performing Select * query.  If number string query as number, if date string query as date.

'http://support.microsoft.com/default.aspx?scid=KB;EN-US;230501
'        NOTE: The Jet OLEDB:Engine Type=4 is only for Jet 3.x format MDB files. If this value is left out, the database is automatically upgraded to the 4.0 version (Jet OLEDB:Engine Type=5). See the following table for appropriate values for Jet OLEDB:Engine Type:
'        Jet OLEDB:Engine Type   Jet x.x Format MDB Files
'1           JET10
'2           JET11
'3           JET2X
'4           JET3X
'5           JET4X

'!!!!! Warning !!!!!:  Even when IMEX is set to 1, values may still return dynamically as different types if registry key above is unmodified.  This should be considered when setting values as well as when setting WHERE clauses as values may not always return as the same type from different workbooks(with the same formatting).
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Example Implementation:''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''ADODB.ObjectStateEnum
'Private Enum ObjectStateEnum
'    adStateClosed = 0
'    adStateOpen = 1
'End Enum
'
'Sub ConnectToDataSource()
'    Dim strDBHost As String
'    Dim strDBName As String
'    Dim strRet As String
'    Dim i As Long
'
'    Dim objDB As clsDBConnectivity
'    Set objDB = New clsDBConnectivity

'
'    'Workbook path
'    strDBHost = objDB.GetWorkbookProperty(ThisWorkbook, "DataSource", True)
'    'Worksheet name
'    strDBName = objDB.GetWorkbookProperty(ThisWorkbook, "ExcelDBTable", True)
'
'    If objDB.ConnectToDB(strDBHost, strDBName, True, True) Then 'If connected
'        Dim oRecordSet As Object 'ADODB.Recordset
'
'        'Workbook path
'        Call objDB.SetWorkbookProperty(ThisWorkbook, "ExcelCapacityDBName", objDB.DataSource)
'        'Worksheet name
'        Call objDB.SetWorkbookProperty(ThisWorkbook, "ExcelCapacityDBTable", objDB.ExcelDBTable)
'
'        Set oRecordSet = CreateObject("ADODB.Recordset")
'        Set oRecordSet = objDB.SelectAllRecordsRecordset(True)
'
'        Do While Not oRecordSet.EOF
'    '        oRecordSet ("<Field/Column Header>")
''            Debug.Print oRecordSet(0).Value & vbTab & oRecordSet(1).Value & vbTab & oRecordSet(2).Value & vbTab & oRecordSet(3).Value '& vbTab & oRecordSet(4).Value & vbTab & oRecordSet(5).Value & vbTab & oRecordSet(6).Value & vbTab & oRecordSet(7).Value & vbTab & oRecordSet(8).Value
'            For i = 0 To oRecordSet.Fields.Count - 1
'                If i = 0 Then
'                    strRet = oRecordSet(i).Value
'                Else
'                    strRet = strRet & vbTab & oRecordSet(i).Value
'                End If
'            Next
'
'            Debug.Print strRet
'            strRet = vbNullString
'            oRecordSet.MoveNext
'        Loop
'
'        If oRecordSet.State = ExcelObjectStateEnum.adStateOpen Then
'            oRecordSet.Close
'        End If
'
'        Call MsgBox(objDB.ExecuteRecordCount)
'    Else
'        Call MsgBox("Error connecting to specified data tables in:" & vbCrLf & "Database: " & strDBHost & vbCrLf & "Datasource: " & strDBName, vbExclamation, "Error")
''        Set objDB = Nothing
'    End If
'
'    Set objDB = Nothing
'End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'MSADO Constants
' ADODB.GetRowsOptionEnum
Private Const adGetRowsRest = -1

'ADODB.ConnectionModeEnum
Private Enum ConnectionModeEnum
    adModeRead = 1  'Not supported for xls
    adModeReadWrite = 3
    adModeRecursive = 4194304
    adModeShareDenyNone = 16 ' Default
    adModeShareDenyRead = 4
    adModeShareDenyWrite = 8
    adModeShareExclusive = 12
    adModeUnknown = 0
    adModeWrite = 2
End Enum

'ADODB.BookmarkEnum
Private Enum BookmarkEnum
    adBookmarkCurrent = 0
    adBookmarkFirst = 1
    adBookmarkLast = 2
End Enum

'ADODB.PersistFormatEnum
Public Enum PersistFormatEnum
    adPersistADTG = 0
    adPersistXML = 1
End Enum

'ADODB.ObjectStateEnum
Public Enum ObjectStateEnum
    adStateClosed = 0
    adStateOpen = 1
End Enum

'ADODB.CursorLocationEnum
Private Enum CursorLocationEnum
    adUseServer = 2
    adUseClient = 3
End Enum

'ADODB.CursorTypeEnum
Private Enum CursorTypeEnum
    adOpenUnspecified = -1  'Is set to adOpenForwardOnly
    adOpenForwardOnly = 0   'Default Most Performant Option - This cursor type is the same as the Static Cursor above, except that it only provides forward movement. If you only need to make a single pass through the recordset, the cursor type will increase performance.
    adOpenKeyset = 1        'It is similar to the Dynamic Cursor above except hat it does not allow access to records added by other users. It allows all types of movements through the recordset.
    adOpenDynamic = 2       'This type of cursor provides a fully updateable recordset. All changes (update, adds, deletions) made by other users while the recordset is open are visible. It also allows all types of movements through the record set.
    adOpenStatic = 3        'It provides a static non-updateable set of records. You should use this recordset type when you’re simply plan to search the recordset. All types of movements are possible with this recordset type. None of the changes made by another user are available to the open recordset.
End Enum

'ADODB.LockTypeEnum
Private Enum LockTypeEnum
    adLockUnspecified = -1
    adLockReadOnly = 1      'Default
    adLockPessimistic = 2
    adLockOptimistic = 3
    adLockBatchOptimistic = 4
End Enum

'ADODB.PromptEnum Values
Private Enum SQLADOConnectionPromptEnum
    adPromptAlways = 1
    adPromptComplete = 2
    adPromptCompleteRequired = 3
    adPromptNever = 4
End Enum

'ADODB.CommandTypeEnum
Public Enum CommandTypeEnum
'    adCmdUnspecified = -1
    adCmdText = 1               'Use to pass query direct.
'    adCmdTable = 2
    adCmdStoredProc = 4     'Use to call stored proc
'    adCmdUnknown = 8
'    adCmdFile = 256
'    adCmdTableDirect = 512 'Use with Seek
End Enum

'ADODB.ParameterDirectionEnum
Public Enum ParameterDirectionEnum
    adParamInput = 1
    adParamOutput = 2
End Enum

'ADODB.ExecuteOptionEnum
Public Enum ExecuteOptionEnum
    adOptionUnspecified = -1
'    adAsyncExecute = 16
'    adAsyncFetch = 32
'    adAsyncFetchNonBlocking = 64
    adExecuteNoRecords = 128
    adExecuteStream = 1024
End Enum

'ADODB.SchemaEnum
Private Enum SchemaEnum
'    adSchemaCatalogs = 1
    adSchemaColumns = 4
    adSchemaTables = 20
End Enum

'ADODB.DataTypeEnum
Public Enum DataTypeEnum
'    adEmpty = 0
    adSmallInt = 2
    adInteger = 3
'    adSingle = 4
    adDouble = 5
'    adCurrency = 6
'    adDate = 7
    adBoolean = 11
'    adVariant = 12
    adTinyInt = 16
    adChar = 129
    adWChar = 130
'    adDBDate = 133
'    adDBTime = 134
    adDBTimeStamp = 135
    adVarChar = 200
    adLongVarChar = 201
    adVarWChar = 202
    adLongVarWChar = 203
End Enum

Public Enum ClearDestinationType
    NoAction = 0
    ClearAll 'Removes values, not formating.
    ClearUsedRangeBelowHeader
    ClearColumnsBelowHeader
    ClearFirstRow
    DeleteAll 'All rows & columns, values and formating.
    DeleteUsedRangeBelowHeader
End Enum

Public Enum PathParseMode
    Path
    FileName
    FileExtension
    FileNameWithoutExtension
End Enum

Public Enum eDataProvider
    Unknown = 0
    FileBased
    DataSourceName 'DSN 'Created in ODBC Data Source Administrator - Administrative Tools\Data Sources (ODBC).  Both User DSN and System DSN work.
    MSSQL2000 ' '7.0
    MSSQLExpress2005
    'Not supported with MDAC 2.8 SP1, however works on MDAC 6.1 (Windows 7)
    MSSQLExpress2005Instance 'Used to connect to User Instance - such as Portfolio Simulator 2011 DB 'http://msdn.microsoft.com/en-us/library/ms254504.aspx
    MSSQL2005
    MSSQL2008
    
    MSOracle
    Oracle
    Composite45 'Compositesw.com
    Composite50 'Compositesw.com
    PostgreSQL
End Enum

Private Enum eDataFileProvider
    NA = 0
    CSVFile '6 'Excel.XlFileFormat.xlCSV
    XMLFile 'Not fully supported.
    
    MSExcel2003XLS '56 'Excel.XlFileFormat.xlExcel8 'Not sure if this is correct number.
    MSExcel2007XLSX
    MSExcel2007XLSM
    MSExcel2007XLSB
    
    MSAccess2003
    MSAccess2007
    
    MSProject2003SP3 'Requires MS Project 2003 SP3 to be installed.
End Enum

Private Enum eDataProviderCategory
    Filebase
    DataSourceName
    database
End Enum

Public Enum eConnectionMode
    SelectMode = 0
    UpdateMode = 1
'    InsertMode = 2
End Enum

Private Enum ODBCRequestFlags
    ODBC_ADD_DSN = 1            'Add user data source
    ODBC_CONFIG_DSN = 2         'Configure (edit) data source
    ODBC_REMOVE_DSN = 3         'Remove data source
    ODBC_ADD_SYS_DSN = 4        'Add system data source
    ODBC_CONFIG_SYS_DSN = 5     'Modify an existing system data source.
    ODBC_REMOVE_SYS_DSN = 6     'Remove an existing system data source.
    ODBC_REMOVE_DEFAULT_DSN = 7 'Remove the default data source specification section from the system information.
End Enum

Private Type OSVERSIONINFO
    OSVSize         As Long         'size, in bytes, of this data structure
    dwVerMajor      As Long         'ie NT 3.51, dwVerMajor = 3; NT 4.0, dwVerMajor = 4.
    dwVerMinor      As Long         'ie NT 3.51, dwVerMinor = 51; NT 4.0, dwVerMinor= 0.
    dwBuildNumber   As Long         'NT: build number of the OS 'Win9x: build number of the OS in low-order word. High-order word contains major & minor ver nos.
    PlatformID      As Long         'Identifies the operating system platform.
    szCSDVersion    As String * 128 'NT: string, such as "Service Pack 3" 'Win9x: string providing arbitrary additional information
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Const UNIQUE_NAME = &H0
Private Const MAX_PATH = 260

Private Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame" 'For VBA UserForm class name

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As Any) As Long
Private Declare Function WritePrivateProfileStringA Lib "kernel32" (ByVal strSection As String, ByVal strKey As String, ByVal strString As String, ByVal strFileNameName As String) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32.dll" (rclsid As GUID) As Long
Private Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hWndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long

'Private Declare Function GetTempPathA Lib "kernel32" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
   
Private m_ExecuteRecordCount As Long
Private m_ExecuteRowsEffectedCount As Long 'Variant
Private m_DBDataSource As String 'Source
Private m_DBDataSourceBackupExcelFilePath As String 'Backup Excel file path used to work around memory leak when performing query on self.
Private m_DBDataFileSource As String 'Workbook file name or Host Name
Private m_DBDataSourceCategory As eDataProviderCategory
Private m_DBDatabase As String
Private m_DBSchema As String
Private m_DBTable As String 'Worksheet name
Private m_TrustedConnection As Boolean
Private m_DBDBSourceType As eDataProvider
Private m_IsExcelFile As Boolean
'Private m_DBDataExcelFileCategory As ExcelDataSourceType
Private m_RecordSetFields As Collection
Private m_Application As Object 'Application
Private m_DBConnection As Object 'ADODB.Connection  'Connection object to the database
Private m_SchemaIniPath As String
Private m_Command As Object 'ADODB.Command   'Used to call stored procs with parameters.  All other calls should be accomplished with oConnection.Execute

'=======================================================================================================
'Properties
'=======================================================================================================
Public Property Get DBConnection() As Object
    'Workbook file name or Host Name
    Set DBConnection = m_DBConnection
End Property

Public Property Get DBDataSource() As String
    'Workbook file name or Host Name
    DBDataSource = m_DBDataSource
End Property

Private Property Let DBDataSource(ByVal strDBSource As String)
    m_DBDataSource = strDBSource
    
    If DBDataSourceType = eDataProvider.FileBased Then
        DBDataFileSourceType = GetDBDataFileTypeByExtension(strDBSource)
    End If
End Property

Private Property Get DataSourceBackupExcelFilePath() As String
    DataSourceBackupExcelFilePath = m_DBDataSourceBackupExcelFilePath
End Property

Private Property Let DataSourceBackupExcelFilePath(ByVal strFilePath As String)
    'Delete previous file if one exists
    Call DeleteTempFile(m_DBDataSourceBackupExcelFilePath)
    
    strFilePath = SaveTempExcelFile(strFilePath)
    If strFilePath <> vbNullString Then
        m_DBDataSourceBackupExcelFilePath = strFilePath
    End If
End Property

Private Property Get SchemaINIPath() As String
    SchemaINIPath = m_SchemaIniPath
End Property

Private Property Let SchemaINIPath(ByVal strSchemaINI As String)
    m_SchemaIniPath = strSchemaINI
End Property

Private Property Get DBDataFileSourceType() As eDataFileProvider
    'Workbook file name or Host Name
    DBDataFileSourceType = m_DBDataFileSource
End Property

Private Property Let DBDataFileSourceType(ByVal DBFileSourceType As eDataFileProvider)
    m_DBDataFileSource = DBFileSourceType
    DBDataFileIsExcel = GetDataFileIsExcel(m_DBDataFileSource)
End Property

Private Property Get DBDataFileIsExcel() As Boolean
    DBDataFileIsExcel = m_IsExcelFile
End Property

Private Property Let DBDataFileIsExcel(IsExcel As Boolean)
    m_IsExcelFile = IsExcel
End Property

Private Property Get DBDataSourceCategory() As eDataProviderCategory
    DBDataSourceCategory = m_DBDataSourceCategory
End Property

Private Property Let DBDataSourceCategory(ByVal DataSourceCategory As eDataProviderCategory)
    m_DBDataSourceCategory = DataSourceCategory
End Property

Private Property Get DBDatabase() As String
    DBDatabase = m_DBDatabase
End Property

Private Property Let DBDatabase(ByVal DBDatabase As String)
    m_DBDatabase = DBDatabase
End Property

Private Property Get DBTable() As String
    DBTable = m_DBTable
End Property

Private Property Let DBTable(ByVal strDBTable As String)
    'GBH-Fix: Verify active connection then verify if valid against list before setting?
    m_DBTable = strDBTable
End Property

Private Property Get DBSchema() As String
    DBSchema = m_DBSchema
End Property

Private Property Let DBSchema(ByVal strDBSchema As String)
    m_DBSchema = strDBSchema
End Property

Private Property Get DBTrustedConnection() As Boolean
    DBTrustedConnection = m_TrustedConnection
End Property

Private Property Let DBTrustedConnection(ByVal bTrustedConnection As Boolean)
    m_TrustedConnection = bTrustedConnection
End Property

Public Property Get DBQueryTableName() As String
    Select Case DBDataSourceCategory 'DBDataSourceType
        'Files
        Case eDataProviderCategory.Filebase
            DBQueryTableName = "[" & DBTable & "]"
        'DB
        Case eDataProviderCategory.database
            Select Case DBDataSourceType
                Case eDataProvider.MSSQL2000, eDataProvider.MSSQLExpress2005, eDataProvider.MSSQLExpress2005Instance, eDataProvider.MSSQL2005, eDataProvider.Composite45, eDataProvider.Composite50, eDataProvider.PostgreSQL
                    DBQueryTableName = "[" & DBDatabase & "].[" & DBTable & "]"
                Case eDataProvider.MSOracle, eDataProvider.Oracle
                    DBQueryTableName = Chr(34) & DBTable & Chr(34)
            End Select
        Case eDataProviderCategory.DataSourceName
            Debug.Assert False 'Untested
            DBQueryTableName = DBTable
    End Select
End Property

Public Function DBQueryFormattedDateString(OriginalDate As Date) As String
'Date formatting notes:
    'ODBC Date Literal: http://msdn.microsoft.com/en-us/library/ms187819.aspx
'            strQuery = " AND (PatientRecruitData.CDS_DATE <= {ts '" & Format(Now(), "yyyy-mm-dd hh:mm:ss") & "'})"

    Select Case DBDataSourceCategory 'DBDataSourceType
        'Files
        Case eDataProviderCategory.Filebase 'True for CSV, Excel2003, and possibly others.
        'Excel column needs to be formatted as Date #m/d/yyyy#)"
        'When querying based on dates, format as #3/13/08# such as : [CDS_DATE] <= #" & Format(Now(), "m/d/yyyy;@") & "#"
            DBQueryFormattedDateString = "#" & Format(OriginalDate, "m/d/yyyy") & "#"
        'DB
        Case eDataProviderCategory.database
            Select Case DBDataSourceType
'                Case eDataProvider.MSSQL2000, eDataProvider.MSSQLExpress2005, eDataProvider.MSSQLExpress2005Instance, eDataProvider.MSSQL2005, eDataProvider.MSSQL2008, eDataProvider.Composite45
'                    DBQueryTableName = "[" & DBDatabase & "].[" & DBTable & "]"
                Case eDataProvider.MSOracle, eDataProvider.Oracle
                    'http://www.experts-exchange.com/Programming/Languages/Visual_Basic/VB_DB/Q_22779759.html
                    'I also believe it's a date format problem.  In Oracle it's always best to convert strings to dates yourself and not let Oracle do it for you.
                    'SELECT * from table where date<=to_date('22-Aug-2007','DD-Mon-YYYY')
                    'You also need to keep in mind the time portion of Oracle dates.  By default to_date sets the timestampe to 00:00:00 so any record added after midnight wouldn't be returned.
                       
                    'The 3 most common ways around this:
                    '1: trunc the date column
                    'SELECT * from table where trunc(date)<=to_date('22-Aug-2007','DD-Mon-YYYY');
                    'problem here is any index on the 'date' column will be ignored.
                    '2: add the time portion
                    'SELECT * from table where date<=to_date('22-Aug-2007 23:59:59','DD-Mon-YYYY HH24:MI:SS');
                    '3: look for 'LESS THAN' the next day at midnight
                    'SELECT * from table where date<to_date('23-Aug-2007','DD-Mon-YYYY');
                    
                    DBQueryFormattedDateString = "TO_DATE('" & Format(OriginalDate, "MM/DD/YYYY") & "', 'MM/DD/YYYY')"
'                    DBQueryFormattedDateString = "TO_DATE('" & Format(OriginalDate, "mm/dd/yyyy;@") & "', 'MM/DD/YYYY')"
            End Select
        Case eDataProviderCategory.DataSourceName
            Debug.Assert False 'Untested
'            DBQueryTableName =
    End Select
End Function

Public Property Get ExecuteRecordCount() As Long
    ExecuteRecordCount = m_ExecuteRecordCount
End Property

Public Property Get ExecuteRowsEffectedCount() As Long
    ExecuteRowsEffectedCount = m_ExecuteRowsEffectedCount
End Property

Private Property Let ExecuteRowsEffectedCount(ByVal Count As Long)
    m_ExecuteRowsEffectedCount = Count
End Property

Private Property Get DBDataSourceType() As eDataProvider
    DBDataSourceType = m_DBDBSourceType
End Property

Private Property Let DBDataSourceType(ByVal DBDataSourceType As eDataProvider)
    m_DBDBSourceType = DBDataSourceType
    DBDataSourceCategory = GetDataSourceCategory(DBDataSourceType)
End Property

Private Property Let RecordSetFields(ByRef colRecordSetFields As Collection)
    Set m_RecordSetFields = Nothing 'Need to free previous?
    Set m_RecordSetFields = New Collection
    Set m_RecordSetFields = colRecordSetFields
End Property

Public Property Get RecordSetFieldsCollection() As Collection
'Returns collection of column field headers associated with last Execute()
    Set RecordSetFieldsCollection = m_RecordSetFields
End Property

'Public Property Get RecordSetFieldsArray() As Variant
''Returns array of column field headers associated with last Execute()
'    RecordSetFieldsArray = CollectionToArray(m_RecordSetFields)
'End Property

Public Property Get RecordSetFieldsArray(Optional IncludeEmptyExcelHeaderCells As Boolean = False) As Variant
'Assumes that query (such as SelectTop1Array()) is performed previous in order to fill collection
    Dim lCount As Long
    Dim vArray As Variant
    Dim vItem As Variant
    Dim col As Collection
    
    Set col = RecordSetFieldsCollection
    
    If Not col Is Nothing Then
        ReDim vArray(col.Count - 1) As Variant
        
        For Each vItem In col
            If DBDataFileIsExcel = True And IncludeEmptyExcelHeaderCells = False And vItem Like ("F#*") Then
                'Don't add it
            Else
                vArray(lCount) = vItem
                lCount = lCount + 1
            End If
        Next
        
        If lCount > 0 Then 'Something was added
            ReDim Preserve vArray(0 To lCount - 1)
            RecordSetFieldsArray = vArray
        End If
    End If
End Property

Public Property Get DBIsConnected() As Boolean
    If Not m_DBConnection Is Nothing Then
        DBIsConnected = (m_DBConnection.State = ObjectStateEnum.adStateOpen)
    End If
End Property

'=======================================================================================================
'Methods
'=======================================================================================================
'Superseded with ExecuteCommand()
'Public Function Execute(strCommandText As String, Optional lTimeout As Long = 30) As Boolean
''Used for calls that return no records.
'    On Error GoTo ErrSub
'
'    Dim lCursor As Long
'
'    ExecuteRowsEffectedCount = -1
'
''    If m_Application.Name = "Microsoft Excel" Then
''        lCursor = m_Application.Cursor
''    End If
'
'    m_DBConnection.CommandTimeout = lTimeout
'    Call m_DBConnection.Execute(strCommandText, ExecuteRowsEffectedCount, ExecuteOptionEnum.adExecuteNoRecords)  'Faster than RecordSetOpen
'    Execute = True
'
'ErrSub: 'Fall through intentional
''    If m_Application.Name = "Microsoft Excel" Then
''        m_Application.Cursor = lCursor
''    End If
'    If Err.Number <> 0 Then
'        Debug.Print Err.Description
'        Err.Raise Err.Number
'    End If
'End Function

Public Function ExecuteArray(strCommandText As String, Optional bStoreRecordCount As Boolean = False, Optional lOptions As ExecuteOptionEnum = ExecuteOptionEnum.adOptionUnspecified, Optional lTimeout As Long = 30, Optional returnFields As Variant) As Variant ', [RecordsAffected], [Options As Long = -1]) As Recordset
'Returns Array and sets m_ExecuteRecordCount if bStoreRecordCount = true to count of returned records.
'*MUCH!* *MUCH!* faster to return array and write array than to loop through recordset, however may be limited in size.  See: http://www.avdf.com/%5Capr98%5Cart_ot003.html
'Pass in returnFields example: returnFields = Array("Collections", "SPE", "ATYA", "Tetanus")
'Check return with IsEmpty(vRecordSet) - will return empty if no records returned.
'Call with Options set to: adExecuteNoRecords for stored procs that return no records.
'SQLExecuteOptionEnum = -1 = adOptionUnspecified

'Example of how to quickly write resultant array contents to sheet.
'shtTest.Cells(1, 1).Resize(UBound(vRecordset, 1) + 1, UBound(vRecordset, 2) + 1) = vRecordset
'shtTest.Cells(1, 1).Resize(UBound(vRecordset, 2) + 1, UBound(vRecordset, 1) + 1) = Application.WorksheetFunction.Transpose(vRecordset)
            
    On Error GoTo errsub
    Dim rsArray As Variant
    Dim lCursor As Long
    Dim i As Long
    Dim colFields As Collection
    Dim Field As Object
    Set colFields = New Collection
'    Dim RecordsAffected As Variant
    Dim oRecordSet As Object    'ADODB.Recordset
                    
    ExecuteRowsEffectedCount = -1
    
    If m_Application.name = "Microsoft Excel" Then
        lCursor = m_Application.Cursor
'        m_Application.Cursor = m_Application.XlMousePointer.xlWait
    End If
        
    m_DBConnection.CommandTimeout = lTimeout
    Set oRecordSet = m_DBConnection.Execute(strCommandText, ExecuteRowsEffectedCount, lOptions)  'Faster than RecordSetOpen

'    Debug.Print strCommandText
    
    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            'Store field header info
            For Each Field In oRecordSet.Fields
                colFields.Add Field.name 'Can't save object collection as it is an interface, so just name
            Next
            RecordSetFields = colFields
            
            If Not oRecordSet.EOF Then
                'GetRows() advances the current cursor position in the RecordSet
                If IsMissing(returnFields) Then
                    rsArray = oRecordSet.GetRows(adGetRowsRest, adBookmarkCurrent)
                Else
                    rsArray = oRecordSet.GetRows(adGetRowsRest, adBookmarkCurrent, returnFields)
                End If
                If bStoreRecordCount = True Then
                    m_ExecuteRecordCount = UBound(rsArray, 2) + 1
                Else
                    m_ExecuteRecordCount = -1
                End If
            End If
        End If
        ExecuteArray = rsArray
    End If
    
errsub: 'Fall through intentional
    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            oRecordSet.Close
        End If
        Set oRecordSet = Nothing
    End If
    
    If m_Application.name = "Microsoft Excel" Then
        m_Application.Cursor = lCursor
    End If
    
    If Err.Number <> 0 Then
        Debug.Print Err.Description
        Err.Raise Err.Number, , Err.Description
    End If
End Function

Public Function ExecuteRecordset(strCommandText As String, Optional bStoreRecordCount As Boolean = False, Optional lOptions As ExecuteOptionEnum = ExecuteOptionEnum.adOptionUnspecified, Optional lTimeout As Long = 30) As Object ', [RecordsAffected], [Options As Long = -1]) As Recordset
'Returns RecordSet and sets m_ExecuteRecordCount if bStoreRecordCount = true to count of returned records.
'Call with Options set to: adExecuteNoRecords for stored procs that return no records.
'SQLExecuteOptionEnum = -1 = adOptionUnspecified
'Calling:
'No return RecordSet:  Execute(strSQL, False, adExecuteNoRecords)
'With return RecordSet: Set oRecordSet = Execute(,strSQL)
    On Error GoTo errsub
    Dim rsArray
    Dim lCursor As Long
    Dim colFields As Collection
    Dim Field As Object
'    Dim RecordsAffected As Variant
    Dim oRecordSet As Object    'ADODB.Recordset

    ExecuteRowsEffectedCount = -1

    If m_Application.name = "Microsoft Excel" Then
        lCursor = m_Application.Cursor
'        m_Application.Cursor = Excel.XlMousePointer.xlWait
    End If
'    lCursor = Application.Cursor
'    Application.Cursor = xlWait

    m_DBConnection.CommandTimeout = lTimeout
    
    Set oRecordSet = m_DBConnection.Execute(strCommandText, ExecuteRowsEffectedCount, lOptions)  'Faster than RecordSetOpen, however Recordset had CursorType of type adOpenForwardOnly, which doesn't allow for moving of cursor.

'    Debug.Print strCommandText

    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            If Not oRecordSet.EOF Then
                'Store field header info
                If Not oRecordSet Is Nothing Then
                    Set colFields = New Collection
                    For Each Field In oRecordSet.Fields
                        colFields.Add Field.name 'Can't save object collection as it is an interface, so just name
                    Next
                    RecordSetFields = colFields
                    Set colFields = Nothing
                End If
                
                If bStoreRecordCount = True Then
                    rsArray = oRecordSet.GetRows() 'GetRows() advances the current cursor position in the RecordSet
                    oRecordSet.MoveFirst 'Can cause requery (slower performance, so optional) 'Recordset.Requery 'May not be able to move cursor if Curortype is adOpenForwardOnly.
                    m_ExecuteRecordCount = UBound(rsArray, 2) + 1
                Else
                    m_ExecuteRecordCount = -1
                End If
            Else
                Set oRecordSet = Nothing
            End If
        Else 'No results returned.
            Set oRecordSet = Nothing
        End If
        Set ExecuteRecordset = oRecordSet
    End If
    
    If m_Application.name = "Microsoft Excel" Then
        m_Application.Cursor = lCursor
    End If
'    Application.Cursor = lCursor

    Exit Function
errsub:
    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            oRecordSet.Close
        End If
        Set oRecordSet = Nothing
    End If
    
    If m_Application.name = "Microsoft Excel" Then
        m_Application.Cursor = lCursor
    End If
    '    Application.Cursor = lCursor

    If Err.Number <> 0 Then
        Debug.Print Err.Description
        Err.Raise Err.Number
    End If
End Function

Public Function SetCommandParameter(ByVal name As String, ByVal DataType As DataTypeEnum, ByVal Direction As ParameterDirectionEnum, ByRef Value As Variant, Optional bInitialize As Boolean = False) As Boolean
'Method to append paramter to use when calling stored Proc via ExecuteStoredProc()
On Error GoTo errsub
    If bInitialize Then
        Call ClearCommandParameters
    End If
    
    m_Command.Parameters.Append CreateParam(name, DataType, Direction, Value)
    SetCommandParameter = True
    Exit Function
errsub:
    SetCommandParameter = False
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Public Sub ClearCommandParameters()
    Do Until m_Command.Parameters.Count = 0
        Call m_Command.Parameters.Delete(0)
    Loop
End Sub

Public Function ExecuteCommand(strCommand As String, Optional TimeOut As Integer = 30) As Object 'm_Command.Parameters
'Returns either cmd.parameters or Recordset Object.
    Set ExecuteCommand = ExecuteDBCommand(CommandTypeEnum.adCmdText, strCommand, False, TimeOut)
End Function

Public Function ExecuteStoredProc(procName As String, Optional ByVal ReturnRecordset As Boolean = False, Optional TimeOut As Integer = 30) As Object 'ADODB.Recordset
'Returns either cmd.parameters or Recordset Object.
    Set ExecuteStoredProc = ExecuteDBCommand(CommandTypeEnum.adCmdStoredProc, procName, ReturnRecordset, TimeOut)
End Function

Private Function ExecuteDBCommand(CommandType As CommandTypeEnum, strCommand As String, ByVal ReturnRecordset As Boolean, Optional TimeOut As Integer = 60) As Object 'ADODB.Recordset
'Public Function CallStoredProc(ByRef cmd As Object, CommandType As CommandTypeEnum, strCommand As String, ByVal ReturnRecordset As Boolean, Optional TimeOut As Integer = 60) As Variant 'ADODB.Recordset
'Returns either cmd.parameters or Recordset Object.
'Command.Execute http://msdn.microsoft.com/en-us/library/windows/desktop/ms681559(v=vs.85).aspx
    On Error GoTo errsub
    
    If DBIsConnected = True Then
        m_Command.ActiveConnection = m_DBConnection     'use the open connection (if none...error!)
        m_Command.CommandType = CommandType             'tell the command we are using a stored procedure or text
        m_Command.CommandText = strCommand              'which proc or text to execute
        m_Command.CommandTimeout = TimeOut
        
        'execute the command and return a recordset or return the params collection
        If ReturnRecordset = True Then
            Set ExecuteDBCommand = m_Command.Execute(m_ExecuteRowsEffectedCount, , ExecuteOptionEnum.adOptionUnspecified)
            m_ExecuteRecordCount = ExecuteDBCommand.RecordCount()
'            Dim oRecordSet As Object    'ADODB.Recordset
'            Set oRecordSet = CreateObject("ADODB.Recordset")
'
'            oRecordSet.CursorLocation = adUseClient
'            oRecordSet.Open m_Command, , adOpenStatic
'
'            'return the recordset
'            Set ExecuteDBCommand = oRecordSet
'
'            'release the recordset
'            If oRecordSet.State = adStateOpen Then
'                oRecordSet.Close
'            End If
'
'            Set oRecordSet = Nothing
        Else  'Return parameters collection.
            Call m_Command.Execute(m_ExecuteRowsEffectedCount, , ExecuteOptionEnum.adExecuteNoRecords)
            Set ExecuteDBCommand = m_Command.Parameters
        End If
    Else
        Debug.Print "Not connected"
        'Debug.Assert False 'Not connected
    End If
    
errsub: 'Fall through intentional
'    If m_Application.Name = "Microsoft Excel" Then
'        m_Application.Cursor = lCursor
'    End If
    Call ClearCommandParameters
    If Err.Number <> 0 Then
'        If Not oRecordSet Is Nothing Then
'            If oRecordSet.State = ObjectStateEnum.adStateOpen Then
'                oRecordSet.Close
'            End If
'            Set oRecordSet = Nothing
'        End If

        'ADO Errors collection - collection can contain items even when connection is working.
        Dim oError As Object 'Error
        For Each oError In m_DBConnection.Errors
            With oError
                Debug.Print "ADO Error#:" & .Number & " Description:" & .Description & " Source:" & .Source
            End With
        Next
        m_DBConnection.Errors.Clear
        
        'VB Error Object
        Debug.Print Err.Description
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End If
End Function

Private Function CreateParam(ByVal name As String, ByVal dtype As DataTypeEnum, ByVal Direction As ParameterDirectionEnum, ByRef Value As Variant) As Object
    'http://www.carlprothman.net/Default.aspx?tabid=97
    'http://www.frentonline.com/Knowledgebase/Development/SQLExpress/Datatype/tabid/362/Default.aspx
    'Used to create a parameter collection for calling a stored proc
    'Allows for passing Null values to DB. Passing "Empty" for NULL values may also work.

    Dim oParams As Object  'ADODB.Parameter
    Set oParams = CreateObject("ADODB.Parameter")

    With oParams
        .name = name
        .Type = dtype
        .Direction = Direction
        Select Case dtype
            Case adInteger  '4 Bytes Long
                .SIZE = 4
                If VBA.IsNull(Value) Then
                    .Value = Null
                Else
                    .Value = CInt(Value)
                End If
            Case adVarChar, adVarWChar, adWChar, adLongVarWChar, adLongVarChar, adChar
                .SIZE = IIf(Len(Value) = 0, 1, Len(Value))
                .Value = CStr(Value) 'IIf(Len(value) = 0, Null, CStr(value))
            Case adDouble
                .SIZE = 8
                .Value = CDbl(Value)
            Case adBoolean
                .SIZE = 1   '2
                If VBA.IsNull(Value) Then
                    .Value = Null
                Else
                    .Value = CBool(Value)
                End If
            Case adDBTimeStamp 'adDate
                .SIZE = 8
                If VBA.IsNull(Value) Or IsEmpty(Value) Then
                    .Value = Null   'Empty may also work.
                Else
                    .Value = CDate(Value)
                End If
            Case adSmallInt  '2 Bytes
                .SIZE = 2
                .Value = CInt(Value)
            Case adTinyInt  '1 Byte
                .SIZE = 1
                .Value = CByte(Value)
            Case Else
                Debug.Print "Internal Error:  Need to add another type!"
                .SIZE = 1
                .Value = Value
        End Select
    End With
    Set CreateParam = oParams
End Function

Public Function DetachDBSYS_SP(DBName As String) As Boolean
'Detatch DB - system stored procedure requires admin rights
    On Error GoTo errsub
    
'    If DoesFilePathExist(DBName) = True And DBName <> vbNullString Then
        If IsSysAdmin Then
            If DropConnections(DBName) = True Then
                Call SetCommandParameter("dbname", adVarChar, adParamInput, DBName)
                Call SetCommandParameter("skipchecks", adVarChar, adParamInput, True)
                Call SetCommandParameter("keepfulltextindexfile", adVarChar, adParamInput, True)
                Call ExecuteStoredProc("sp_detach_db") 'Will error if doesn't exist
                DetachDBSYS_SP = True
            End If
        Else
            Debug.Print "sysadmin role required to perform this operation."
        End If
'    End If
errsub:
End Function

Private Function DropConnections(DBPath As String) As Boolean
'DBPath = full path to MDF.
'Drop active DB connections, rolling back pending transactions.
    On Error GoTo errsub
    
    Dim strCommand As String
'    If DoesFilePathExist(DBPath) = True And DBPath <> vbNullString Then
        strCommand = "ALTER DATABASE [" & DBPath & "] SET SINGLE_USER WITH ROLLBACK IMMEDIATE"
        Call ExecuteCommand(strCommand)
        DropConnections = True
'    End If
errsub:
End Function

Private Function TakeDBOffline(DBPath As String) As Boolean
'DBPath = full path to MDF.
    On Error GoTo errsub
    
    Dim strCommand As String
    If DoesFilePathExist(DBPath) = True And DBPath <> vbNullString Then
        strCommand = "ALTER DATABASE [" & DBPath & "] SET OFFLINE WITH ROLLBACK IMMEDIATE"
        Call ExecuteCommand(strCommand)
        TakeDBOffline = True
    End If
errsub:
End Function

Public Function IsSysAdmin() As Boolean
'http://msdn.microsoft.com/en-us/library/ms176015.aspx
    On Error GoTo errsub
    Dim strCommandText As String
    Dim oRecordSet As Object 'ADODB.Recordset
    
    If DBIsConnected = True Then
        strCommandText = "SELECT CASE WHEN IS_SRVROLEMEMBER('sysadmin') = 1 THEN 1 ELSE 0 END AS DBPERMISSIONS"
        Set oRecordSet = ExecuteDBCommand(adCmdText, strCommandText, True)
        If Not oRecordSet.EOF Then
            IsSysAdmin = (oRecordSet.Fields("DBPERMISSIONS") = 1)
        End If
    End If
    
errsub: 'Fall through intentional
    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            oRecordSet.Close
        End If
        Set oRecordSet = Nothing
    End If
End Function

Public Function AttachSingleFileDBSYS_SP(DBName As String, DBFilePath As String) As Boolean
'Attach DB - Calls system stored procedure sp_attach_single_file_db to attach *.MDF DB file
'Requires admin rights
    On Error GoTo errsub

'    If DoesFilePathExist(DBFilePath) = true And DBName <> vbNullString Then
        If IsSysAdmin() Then
            Call SetCommandParameter("dbname", adVarChar, adParamInput, DBName) 'nvarchar(128)
            Call SetCommandParameter("physname ", adVarChar, adParamInput, DBFilePath) 'nvarchar(260)
            Call ExecuteStoredProc("sp_attach_single_file_db")
            AttachSingleFileDBSYS_SP = True
        Else
            Debug.Print "sysadmin role required to perform this operation."
        End If
'    End If
errsub:
End Function

Public Function ConnectToDB(Provider As eDataProvider, strDataSource As String, Optional ByRef strDatabase As String = vbNullString, Optional ByRef strDataSchema As String = vbNullString, Optional ByRef strDataTable As String = vbNullString, Optional lPort As Long, Optional TrustedConnection As Boolean = True, Optional strUsername As String, Optional strPassword As String, Optional bFileHasHeader As Boolean = True, Optional bPromptForUseFile As Boolean = False, Optional strPromptDialogTitle As String = "File Path", Optional bAllowValidDataTableSelection As Boolean = False, Optional bADODialogPrompt As Boolean = False, Optional ConnectionMode As eConnectionMode = eConnectionMode.SelectMode) As Boolean
'Public Function ConnectToDB(Provider As eDataProvider, ByRef strDataSource As String, Optional ByRef strDatabase As String, Optional ByRef strDataTable As String, Optional TrustedConnection As Boolean = True, Optional strUserName As String, Optional strPassword As String, Optional bFileHasHeader As Boolean = True, Optional bPromptForUseFile As Boolean = False, Optional strPromptDialogTitle As String = "File Path", Optional bAllowValidDataTableSelection As Boolean = False, Optional bADODialogPrompt As Boolean = False) As Boolean
'strDataSource can get changed within function if passed is invalid.
'AKA Initial Catalog.  strDataTable not required to pass in when file type is csv.
'Don't store Path and Table values internally (in Custom Properties) as multiple objects can be used at the same time.
'Excel uses Jet.OLEDB drivers
'When query to worksheet is desired append "$" to name.

    On Error GoTo errsub
    
    DBDataSourceType = Provider
    DBDatabase = strDatabase
    DBSchema = strDataSchema
    DBTable = strDataTable 'sets formatting. 'This needs to be before "DataSource = strDataSource"
    DBDataSource = strDataSource 'Can get reset inside GetValidDataSource()
    DBTrustedConnection = TrustedConnection
    
    'Make sure starting out closed.
    'GBH or could check to see if this is the same connection and leave it open for speed.
    If m_DBConnection.State = ObjectStateEnum.adStateOpen Then
        m_DBConnection.Close
    End If
    
'=========================================================================================================
'Set connection string =======================================================================================
'www.connectionstrings.com
'=========================================================================================================

    With m_DBConnection
        strDataSource = GetValidDataSource(Provider, strDataSource, strDatabase, bPromptForUseFile, strPromptDialogTitle, TrustedConnection)
        Select Case DBDataSourceCategory
            Case eDataProviderCategory.Filebase
                Select Case DBDataFileSourceType
                    Case eDataFileProvider.CSVFile
                        'http://weblogs.asp.net/fmarguerie/archive/2003/10/01/29964.aspx
                        'Warning: Have seen csv truncate floating point numbers to integers.
                        'Doesn't support UPDATE.
'                        strDataSource = GetValidDataSource(Provider, strDataSource, strDatabase, strDataTable, bPromptForUseFile, strPromptDialogTitle)
                        .Provider = "Microsoft.Jet.OLEDB.4.0;"
                        .Properties("Data Source") = ParsePath(strDataSource, Path)
                        .Properties("Extended Properties") = "text;HDR=" & IIf(bFileHasHeader, "Yes", "No") & ";FMT=CSVDelimited;IMEX=1;MaxScanRows=0" '0 - look at the entire file and set type based on majority.  This can set the wrong value if data is mixed: http://support.microsoft.com/kb/282263
        '                string.Format(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};" + @" Extended Properties=""{1}""", path, "Text;HDR=YES;FMT=Delimited;IMEX=1");
                        'Possibly need to set these:
        '                HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text\ImportMixedTypes = "Text"
        '                HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Text\MaxScanRows = "0"
                        
                        'For CSV, file is not accessible if open in Excel, however they are accessible if open in notepad.  IsFileOpen() only returns true if open in Excel, so this works.
                        If strDataSource = vbNullString Then
                            GoTo errsub
                        Else
                            If IsFileOpen(strDataSource) Then
                                Call MsgBox("Close the following file to proceed." & vbCrLf & vbCrLf & strDataSource, vbInformation + vbOKOnly, "Error")
                                GoTo errsub
                            End If
                        End If
                        
                        'Write out schema.ini file to force driver to import columns as text.
                        If WriteCSVSchemaIni(strDataSource) = False Then
                            GoTo errsub
                        End If
                        
                    'Can also use Microsoft.ACE.OLEDB.12.0 driver to access xls if it is on the computer.
                    Case eDataFileProvider.MSExcel2003XLS ', eDataFileProvider.MSEXCEL2003XLSNamedRange
'                        strDataSource = GetValidDataSource(Provider, strDataSource, strDatabase, strDataTable, bPromptForUseFile, strPromptDialogTitle, TrustedConnection)
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .Properties("Data Source") = strDataSource
'                        .Properties("Mode") = IIf(ConnectionMode = eConnectionMode.SelectMode, ConnectionModeEnum.adModeRead, ConnectionModeEnum.adModeShareDenyNone)
                        .Properties("Extended Properties") = "Excel 8.0;HDR=" & IIf(bFileHasHeader, "Yes", "No") & IIf(ConnectionMode = eConnectionMode.SelectMode, ";IMEX=1;MaxScanRows=0", vbNullString) 'MaxScanRows=8-16 '0 - look at the entire file and set type based on majority.  This can set the wrong value if data is mixed: http://support.microsoft.com/kb/282263
        '                .Properties("User Index") = "Admin"
        '                .Properties("Password") = ""

                        'Possibly need to set these:
        '                HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel\ImportMixedTypes = Text
        '                HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Jet\4.0\Engines\Excel\TypeGuessRows = 0
    
                    Case eDataFileProvider.MSExcel2007XLSX, eDataFileProvider.MSExcel2007XLSM, eDataFileProvider.MSExcel2007XLSB
'                        strDataSource = GetValidDataSource(Provider, strDataSource, strDatabase, strDataTable, bPromptForUseFile, strPromptDialogTitle)
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Data Source") = strDataSource
'                        .Properties("Mode") = ConnectionModeEnum.adModeShareDenyNone 'IIf(ConnectionMode = eConnectionMode.SelectMode, ConnectionModeEnum.adModeRead, ConnectionModeEnum.adModeShareDenyNone)
                        Select Case DBDataFileSourceType
                            Case eDataFileProvider.MSExcel2003XLS, eDataFileProvider.MSExcel2007XLSX    'DBDataFileSourceType = eDataFileProvider.MSEXCEL2007XLSXNamedRange Then 'Office 2007 'Treat all data as text (using "Xml")
                                .Properties("Extended Properties") = "Excel 12.0 Xml;HDR=" & IIf(bFileHasHeader, "Yes", "No") & IIf(ConnectionMode = eConnectionMode.SelectMode, ";IMEX=1;MaxScanRows=0", vbNullString) '0 - look at the entire file and set type based on majority.  This can set the wrong value if data is mixed: http://support.microsoft.com/kb/282263
                            Case eDataFileProvider.MSExcel2007XLSM 'DBDataFileSourceType = eDataFileProvider.MSEXCEL2007XLSMNamedRange Then 'Office 2007 with macros enabled.
                                .Properties("Extended Properties") = "Excel 12.0 Macro;HDR=" & IIf(bFileHasHeader, "Yes", "No") & IIf(ConnectionMode = eConnectionMode.SelectMode, ";IMEX=1;MaxScanRows=0", vbNullString)
                            Case eDataFileProvider.MSExcel2007XLSB 'DBDataFileSourceType = eDataFileProvider.MSEXCEL2007XLSBNamedRange Then 'Office 2007 Open XML file format with macros enabled.
                                .Properties("Extended Properties") = "Excel 12.0;HDR=" & IIf(bFileHasHeader, "Yes", "No") & IIf(ConnectionMode = eConnectionMode.SelectMode, ";IMEX=1;MaxScanRows=0", vbNullString)
                        End Select
                    
                    Case eDataFileProvider.MSAccess2003
                        .Provider = "Microsoft.Jet.OLEDB.4.0"
                        .Properties("Data Source") = strDataSource
        '                .Properties("Initial Catalog") = strDatabase 'strDataTable
                        
                    Case eDataFileProvider.MSAccess2007
                        .Provider = "Microsoft.ACE.OLEDB.12.0"
                        .Properties("Data Source") = strDataSource
'                        .Properties("Jet OLEDB:Database Password") = strPassword
                    
                    Case eDataFileProvider.XMLFile
                        .Provider = "MSDAOSP"  'OLE DB Simple Provider DLL, MDAC 2.7 or later.
                        .Properties("Data Source") = "MSXML2.DSOControl" '.2.6"
                        
                    Case eDataFileProvider.MSProject2003SP3 'Requires MS Project 2003 SP3 to be installed.
                        .Provider = "Microsoft.Project.OLEDB.11.0"
                        .Properties("Data Source") = strDataSource
                    Case Else 'Case eDataFileProvider.NA
                        Debug.Print "Attempting ADO connection to unknown file type failed."
                        GoTo errsub
                End Select
            
            Case eDataProviderCategory.DataSourceName 'DSN 'Created in ODBC Data Source Administrator - Administrative Tools\Data Sources (ODBC).  Both User DSN and System DSN work.
                'http://www.carlprothman.net/Technology/ConnectionStrings/ODBCDSN/tabid/89/Default.aspx
                .Provider = "MSDASQL" 'Default
                .Properties("Data Source") = strDataSource
                If TrustedConnection = False Then
                    .Properties("User ID") = strUsername
                    .Properties("Password") = strPassword
                End If
    
            Case eDataProviderCategory.database
                Select Case DBDataSourceType
                    Case eDataProvider.MSSQL2000, eDataProvider.MSSQLExpress2005, eDataProvider.MSSQLExpress2005Instance, eDataProvider.MSSQL2005, eDataProvider.MSSQL2008
                        If DBDataSourceType = eDataProvider.MSSQLExpress2005 Or DBDataSourceType = eDataProvider.MSSQLExpress2005Instance Then '(NativeClient) - used to connect to SQL Express 2005
                            .Provider = "SQLNCLI"
                        ElseIf DBDataSourceType = eDataProvider.MSSQL2008 Then
                            .Provider = "SQLNCLI10"
                        Else
                            .Provider = "SQLOLEDB"
                        End If
                        
                        .Properties("Data Source") = strDataSource
                        
                        If DBDataSourceType = eDataProvider.MSSQLExpress2005Instance Then
                            'Supported in MDAC 6.1, but not 2.8 when user is not Windows admin.
                            'if Val(m_DBConnection.Version) >= 6.1
                            .Properties("Extended Properties") = "AttachDbFilename=" & strDatabase & ";User Instance=True" ';Database=strDataTable 'The name for the copied database, if not used the file name is used as the name.
                            .Properties("Integrated Security") = "SSPI" 'for OleDB
                            'Integrated Security=True;
                            'Trusted Connection=Yes;
                        Else
                            .Properties("Initial Catalog") = strDatabase
'                            .Properties("Encrypt") = "Yes"
                            
                            If TrustedConnection = True Then
                                .Properties("Integrated Security") = "SSPI" '.Properties("Trusted_Connection=Yes")
                            ElseIf TrustedConnection Then
                                .Properties("User ID") = strUsername '"User Index"
                                .Properties("Password") = strPassword
                            End If
                            '.Properties("Network Library") = "dbmssocn" 'Use TCP/IP 'http://www.asp101.com/articles/john/connstring/default.asp
                        End If
                              
                    Case eDataProvider.Oracle 'Oracle Provider for OLE DB
                    'Oracle driver obtained from:("Oracle 11g ODAC and Oracle Developer Tools for Visual Studio") : http://www.oracle.com/technology/software/tech/dotnet/utilsoft.html
                        .Provider = "OraOLEDB.Oracle"
                        .Properties("Data Source") = strDataSource
                '                .Properties("Data Source") = strDataSource & "/" & strDataTable
                        .Properties("User Index") = strUsername 'SYSMAN
                        .Properties("Password") = strPassword 'orcl
                    
                    Case eDataProvider.MSOracle 'Microsoft OLE DB Provider for Oracle
                        .Provider = "MSDAORA"
                        .Properties("Data Source") = strDataSource
                        .Properties("User Index") = strUsername
                        .Properties("Password") = strPassword
                        
                    Case eDataProvider.Composite45 'ODBC - DSN less
                        .Provider = "MSDASQL"
                        lPort = IIf(IsMissing(lPort) = False, 9501, lPort)
                        .Properties("Extended Properties") = "DRIVER=Composite 4.5;HOST=" & strDataSource & ";PORT=" & lPort & ";UID=" & strUsername & ";PWD=" & strPassword & ";DOMAIN=" & strDatabase & ";DATASOURCE=PDSI;TRACELEVEL=off;CONNECTTIMEOUT=999" ';CATALOG=CTS"
                    
                    Case eDataProvider.Composite50 'ODBC - DSN less
                        .Provider = "MSDASQL"
                        lPort = IIf(IsMissing(lPort) = False, 9411, lPort)
                        .Properties("Extended Properties") = "DRIVER=Composite 5.0;HOST=" & strDataSource & ";PORT=" & lPort & ";UID=" & strUsername & ";PWD=" & strPassword & ";DOMAIN=" & strDatabase & ";DATASOURCE=PDSI;TRACELEVEL=off;CONNECTTIMEOUT=999" ';CATALOG=CTS"
                    
                    Case eDataProvider.PostgreSQL 'ODBC - DSN less
                        'Driver Installation Location: https://www.postgresql.org/ftp/odbc/versions/msi
                        .Provider = "MSDASQL"
                        lPort = IIf(IsMissing(lPort) = False, 5432, lPort)
                        .Properties("Extended Properties") = "DRIVER=PostgreSQL Unicode;Database=" & strDatabase & ";PORT=" & lPort & ";Server=" & strDataSource & ";Uid=" & strUsername & ";Pwd=" & strPassword
                End Select
            Case Else
                MsgBox "Unimplemented connection type"
        End Select
        If strPassword <> vbNullString Then
            .Properties("Persist Security Info") = False 'Default, don't include with eDataProvider.MSSQLExpress2005Instance
        End If
    End With
    
'=========================================================================================================
'=Verify Data Table =========================================================================================
'=========================================================================================================
    If strDataSource <> vbNullString Then 'User may have selected Cancel in browse for file dialog.
        If bADODialogPrompt = True Then
            m_DBConnection.Properties("Prompt") = SQLADOConnectionPromptEnum.adPromptAlways   'Prompt for connection information.
    '        GoTo errsub
        End If
        
'        Debug.Print m_DBConnection
'        Debug.Print m_DBConnection.Properties("Extended Properties")
        m_DBConnection.Open

        If m_DBConnection.State = ObjectStateEnum.adStateOpen Then
            Select Case DBDataSourceCategory
                Case eDataProviderCategory.Filebase
                    Select Case DBDataFileSourceType
                        Case eDataFileProvider.CSVFile
                            DBTable = ParsePath(strDataSource, FileName) 'For CSV, table is file name.
                            ConnectToDB = True
                        Case eDataFileProvider.MSAccess2003, eDataFileProvider.MSAccess2007
                            ConnectToDB = GetValidDataTableName(bAllowValidDataTableSelection)
                        Case eDataFileProvider.MSExcel2003XLS, eDataFileProvider.MSExcel2007XLSX, eDataFileProvider.MSExcel2007XLSM, eDataFileProvider.MSExcel2007XLSB
                            ConnectToDB = GetValidDataTableName(bAllowValidDataTableSelection)
                        Case eDataFileProvider.MSProject2003SP3
                            ConnectToDB = GetValidDataTableName(bAllowValidDataTableSelection)
                    End Select
                    
                Case eDataProviderCategory.database
                    ConnectToDB = GetValidDataTableName(bAllowValidDataTableSelection)
                Case eDataProviderCategory.DataSourceName
                    ConnectToDB = GetValidDataTableName(bAllowValidDataTableSelection)
            End Select
        End If
    End If
    
errsub: 'Fall through intentional
'    MsgBox Replace("Failed to connect to '###' database.", "###", strDataSource) & vbCrLf & "Error: " & Err.Number & ", " & Err.Description, vbCritical + vbOKOnly, "Database Error"
    If Err.Number <> 0 Then
        If Not m_DBConnection Is Nothing Then
            'ADO Errors collection - collection can contain items even when connection is working.
            Dim oError As Object 'Error
            For Each oError In m_DBConnection.Errors
                With oError
                    Debug.Print "ADO Error#:" & .Number & " Description:" & .Description & " Source:" & .Source
                End With
            Next
            m_DBConnection.Errors.Clear
            
            If m_DBConnection.State = ObjectStateEnum.adStateOpen Then
                m_DBConnection.Close
            End If
        End If
        If Err.Number = -2147217842 Then 'Cancel button selected on ValidDataTableSelection dialog.
            'Const adErrOperationCancelled = 3712 (&HE80)
'            Call MsgBox("Operation Canceled", vbOKOnly, "Error")
        Else
            Debug.Print Err.Description
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
        End If
    End If
End Function

Public Function GetISAMName(Provider As eDataProvider, ByRef strDataSource As String, Optional ByRef strDatabase As String, Optional ByRef strDataTable As String, Optional TrustedConnection As Boolean = True, Optional strUsername As String, Optional strPassword As String, Optional bFileHasHeader As Boolean = True) As String
'Used to specify DB syntax for use in place of DB name in query to allow for direct DB to DB data transfer!!!
'If the same connection is used for multiple queries using ISAM names, connection can get corrupted (data will not return correctly).
'Solution is to reconnect for each ISAM query, which is slightly slower, but only by a little.
'Fastest way to transfer data.  Just surround existing query with:
'"SELECT | INSERT * INTO " & objDB.GetISAMName(FileBased, DestinationDB, , DestinationTable) & " FROM (<Existing Query Here>)"
'Common use is: SELECT | INSERT * INTO
'Jet ISAM = Indexed Sequential Access Method
'Can use in query for eaither the source of destination DB
'[<Full path to Microsoft Access database>].[<Table Name>]
'[ODBC;<ODBC Connection String>].[<Table Name>]
'[<ISAM Name>;<ISAM Connection String>].[<Table Name>]
'http://support.microsoft.com/kb/321686
'http://support.microsoft.com/kb/200427
'http://www.mikesdotnetting.com/Article/79/Import-Data-From-Excel-to-Access-with-ASP.NET
'http://msdn.microsoft.com/en-us/library/ms709353.aspx

'[Text;Database=C:\\Path\;HDR=YES].[txt.csv]
'[Excel 8.0;Database=c:\\customers.xls;HDR=Yes].[Sheet1]
'[MS Access;Database=c:\\customers.mdb].[New Table]
'[ODBC;Driver=SQL Server;SERVER=XXX;DATABASE=Pubs;UID=<username>;PWD=<strong password>;].[RemoteShippers]

'strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & "[Text;Database=" & ParsePath(strFilePath, Path) & ";HDR=YES].[" & ParsePath(strFilePath, FileName) & "]"
'strCommand = "SELECT * INTO [Excel 8.0;HDR=Yes;IMEX=1;MaxScanRows=16;Database=" & strFilePath & "].[" & strDataTable & "$] FROM " & "[" & strDataTableDestination & "]"
'strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & "[Excel 8.0;HDR=Yes;IMEX=1;MaxScanRows=16;Database=" & strFilePath & "].[" & strExcelWorksheet & "]"
'strCommand = "SELECT * INTO [MS Access;Database=" & strDBPath & "].[" & strDataTable & "] FROM " & "[" & strExcelWorksheet & "]"
'strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & "[MS Access;Database=" & strFilePath & "].[" & strDataTable & "]"
    Dim ISAMConnectionString As String

    Select Case GetDataSourceCategory(Provider)
        Case eDataProviderCategory.Filebase
            Select Case GetDBDataFileTypeByExtension(strDataSource) 'DBDataFileSourceType
                Case eDataFileProvider.CSVFile
                    ISAMConnectionString = "[Text;Database=" & ParsePath(strDataSource, Path) & ";HDR=" & IIf(bFileHasHeader, "Yes", "No") & "].[" & ParsePath(strDataSource, FileName) & "]"
                Case eDataFileProvider.MSExcel2003XLS, eDataFileProvider.MSExcel2007XLSX, eDataFileProvider.MSExcel2007XLSM, eDataFileProvider.MSExcel2007XLSB
                    'ISAM version 8.0 drivers used for 2003 and 2010 files.  Pass in trailing "$" for worksheet name.
                    ISAMConnectionString = "[Excel 8.0;HDR=" & IIf(bFileHasHeader, "Yes", "No") & ";IMEX=1;MaxScanRows=16;Database=" & strDataSource & "].[" & strDataTable & "]"
                Case eDataFileProvider.MSAccess2003, eDataFileProvider.MSAccess2007
                    ISAMConnectionString = "[MS Access;Database=" & strDataSource & "].[" & strDataTable & "]"
                Case eDataFileProvider.MSProject2003SP3
                    Debug.Assert False 'Unimplemented
'                    ISAMConnectionString = "[MS Project;Database=" & strDataSource & "].[" & strDataTable & "]"
                Case eDataFileProvider.NA
                    Debug.Assert False 'add more?
            End Select
        Case eDataProviderCategory.database
            Select Case Provider
                Case eDataProvider.MSSQL2000, eDataProvider.MSSQLExpress2005, eDataProvider.MSSQLExpress2005Instance, eDataProvider.MSSQL2005, eDataProvider.MSSQL2008
                    ISAMConnectionString = "[ODBC;Driver=SQL Server;SERVER=" & strDataSource & ";DATABASE=" & strDatabase & ";UID=" & strUsername & ";PWD=" & strPassword & ";].[" & strDataTable & "]"
                Case eDataProvider.Unknown
                    Debug.Assert False 'add more?
            End Select
    End Select
'    Debug.Print ISAMConnectionString
    GetISAMName = ISAMConnectionString
End Function

Public Function ConfigureODBCDataSource(Driver As String, Server As String, DSN As String, database As String, Optional Description As String) As Boolean
    'Modifed from http://support.microsoft.com/kb/171146
    'Use ODBC Data Source Administrator to determine settings.
    'If hwndParent is supplied, dialog will show.
    
    Dim lRet As Long
    Dim Attributes As String
    
    'Remove previous instance if there is one.
    Call RemoveODBCDataSource(Driver, DSN)
    
    'Set the driver to SQL Server because it is most common.
    'Set the attributes delimited by null.
    'See driver documentation for a complete list of supported attributes.
    Attributes = "SERVER=" & Server & Chr(0)
    If Description <> vbNullString Then
        Attributes = Attributes & "DESCRIPTION=" & Description & Chr(0)
    End If
    Attributes = Attributes & "DSN=" & DSN & Chr(0)
    Attributes = Attributes & "DATABASE=" & database & Chr(0)
    Attributes = Attributes & "Trusted_Connection=Yes" & VBA.Chr(0) 'Use Windows Authentication
    If SQLConfigDataSource(0, ODBC_ADD_DSN, Driver, Attributes) Then
        ConfigureODBCDataSource = True
    End If
End Function

Public Function RemoveODBCDataSource(Driver As String, DSN As String) As Boolean
    'Modifed from http://support.microsoft.com/kb/171146
    'Set the driver to SQL Server because most common.
    'Set the attributes delimited by null.
    'See driver documentation for a complete list of attributes.
    'If hwndParent is supplied, dialog will show.
    
    If SQLConfigDataSource(0, ODBC_REMOVE_DSN, Driver, "DSN=" & DSN & Chr(0)) Then
        RemoveODBCDataSource = True
    End If
End Function

Public Function IsFieldValid(TestFieldName As String) As Boolean
    'Return if passed in field is member of connected dataset.
    On Error GoTo errsub
    
    Dim oRecordSet As Object 'ADODB.Recordset
    Dim retVal As String
    
    'Test if field is in connected DB
    If DBIsConnected = True Then
        Set oRecordSet = SelectTop1RecordSet
        
        If Not oRecordSet.EOF Then
            retVal = oRecordSet.Fields(TestFieldName)
            IsFieldValid = True
        End If
    End If
    
errsub: 'Fall through intentional
    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            oRecordSet.Close
        End If
        Set oRecordSet = Nothing
    End If
End Function

Private Function GetValidDataSource(DBDataProvider As eDataProvider, ByRef strDataSource As String, ByRef strDataTable As String, Optional bPromptForUseFile As Boolean = False, Optional strPromptDialogTitle As String = "File Path", Optional TrustedConnection As Boolean = True) As String ', Optional TrustedConnection As Boolean = True, Optional DataSourceName As String) As String
'Could simplify and replace with: m_DBConnection.Properties("Prompt") = SQLADOConnectionPromptEnum.adPromptAlways
    Dim strNewFile As String
    Dim strFile As String
    Dim bPreviousSetting As Boolean
    Dim strServerName As String
    Dim strInstanceName As String
    Dim strLocalMachineName As String
    Dim oCommonDlg As clsCommonDlg
    
    Set oCommonDlg = New clsCommonDlg
    
    Select Case DBDataSourceCategory 'DBDataProvider
        'File based
        Case eDataProviderCategory.Filebase
            If m_Application.name = "Microsoft Excel" Then
                bPreviousSetting = m_Application.ScreenUpdating
                If m_Application.ScreenUpdating = False Then
                    m_Application.ScreenUpdating = True
                End If
            End If
            
            'Get export file.
            strNewFile = oCommonDlg.GetValidFile(strDataSource, oCommonDlg.GetFilterFileTypeByExtension(strDataSource), strPromptDialogTitle, bPromptForUseFile)
            
            'GetValidFile() can return empty string
            If strNewFile <> vbNullString And DoesFilePathExist(strNewFile) = True Then
                DBDataSource = strNewFile
                'Work around for memory leak bug in Excel described at: http://support.microsoft.com/kb/319998
                'Copy file to temp folder, which is then queried against to avoid memory leak.
                'Memory leak occurs for all version of Excel (tested 2003, 2007, 2010) and all file types (xls, xlsx, xlsm).
                If m_Application.name = "Microsoft Excel" And DBDataFileIsExcel = True Then
                    strFile = ParsePath(strNewFile, FileName)
                    If StrComp(m_Application.ActiveWorkbook.name, strFile) = 0 Then 'Same
                        DataSourceBackupExcelFilePath = strFile
                        
                        'Reset file type as it may have changed.
                        DBDataSource = DataSourceBackupExcelFilePath
                    End If
                End If
                
                GetValidDataSource = DBDataSource
            End If
            
            If m_Application.name = "Microsoft Excel" Then
                m_Application.ScreenUpdating = bPreviousSetting
            End If
        
        'DB
        Case eDataProviderCategory.database
            Select Case DBDataProvider
                Case eDataProvider.MSSQL2000, eDataProvider.MSSQLExpress2005, eDataProvider.MSSQLExpress2005Instance, eDataProvider.MSSQL2005, eDataProvider.MSSQL2008, eDataProvider.PostgreSQL
'                    If VBA.InStr(1, strDataSource, "\pipe\") = 0 Then 'Make sure it's not a Named Pipe instance
'                        Call GetServerInstanceParts(strDataSource, strServerName, strInstanceName)
'                        strLocalMachineName = GetLocalMachineNameWMI
'                        'Check if local machine
'                        If strServerName = "." Or strServerName = "(LOCAL)" Or UCase(strServerName) = UCase(strLocalMachineName) Then
'                            'Make sure service is running if on local machine.
'                            If StartServiceForTimeWMI(strServerName, "MSSQL$" & strInstanceName, 10) = False Then 'If no instance used don't include "$"
'                                MsgBox "Error starting database service: MSSQL$" & strInstanceName & vbCrLf & "On server: " & strServerName, vbExclamation, "Database Error" 'If no instance used don't include "$"
'                                Exit Function
'                            End If
'
'                            If DBDataProvider = eDataProvider.MSSQLExpress2005Instance Then
'                                GetValidDataSource = strDataSource
'                            ElseIf GetValidSQLDBName(DBDataProvider, strDataSource, strDataTable, TrustedConnection, bPromptForUseFile) = True Then
'                                GetValidDataSource = strDataSource
'                            End If
'                        End If
'                    Else
                        GetValidDataSource = strDataSource
'                    End If
            End Select
        Case eDataProviderCategory.DataSourceName
            GetValidDataSource = strDataSource
    End Select
End Function

Public Function CopyRecordsetToExcelRange(ByRef Recordset As Object, rngDestination As Object, ClearDestination As ClearDestinationType, Optional AutoFit As Boolean = True, Optional AutoFormat As Boolean = False, Optional ResetCursor As Boolean = False) As Object 'Excel.Range
'Copies recordset to Excel sheet given recorset and start cell of range destination. Derived from Excel help for CopyFromRecordset.
'Returns modified range
'Can also use "SELECT * INTO" syntax to accomplish similar. 'http://support.microsoft.com/kb/200427
'Sets RecordSet cursor to the end of the file by default.
'Can cause cells to have numbers stored as text.  To avoid the issue instead use CopyRecordsetArrayToExcelRange()
    On Error GoTo errsub
    
    Const xlUp = -4162
    
    Dim MaxRows As Variant
    Dim MaxColumns As Variant
    Dim bScreenUpdating As Boolean
    Dim bEnableEvents As Boolean
    Dim rng As Object 'Excel.Range
    Dim wksParent As Object 'Excel.Worksheet
    Dim strTemp As String
    Dim lRetVal As Long
    
    Set wksParent = rngDestination.Parent
    Set rngDestination = rngDestination(1)  'Only use first cell.

    bScreenUpdating = m_Application.ScreenUpdating
    bEnableEvents = m_Application.EnableEvents
    MaxRows = (wksParent.Rows.Count - rngDestination.Row) + 1
    MaxColumns = (wksParent.Columns.Count - rngDestination.Column) + 1

    If Not Recordset Is Nothing Then
        If m_Application.ScreenUpdating = True Then
            m_Application.ScreenUpdating = False 'Causes "Filling Cells" status message in Excel 2007.
        End If
        m_Application.EnableEvents = False 'Causes application flicker if on.
        
        Select Case ClearDestination
            Case ClearDestinationType.ClearUsedRangeBelowHeader
                'Clear Range, leaving formatting.
                If Not GetLastUsedRow(wksParent) Is Nothing And Not GetLastUsedColumn(wksParent) Is Nothing Then
                    If rngDestination.Row <= GetLastUsedRow(wksParent).Row And rngDestination.Column <= GetLastUsedColumn(wksParent).Column Then
                        rngDestination.Resize(GetLastUsedRow(wksParent).Row - rngDestination.Row + 1, GetLastUsedColumn(wksParent).Column - rngDestination.Column + 1).ClearContents
                    End If
                End If
            Case ClearDestinationType.ClearColumnsBelowHeader
                'Clear columns in destination area, leaving formatting.
                If Not GetLastUsedRow(wksParent) Is Nothing Then
                    If GetLastUsedRow(wksParent).Row > 0 And Recordset.Fields.Count > 0 Then
                        rngDestination.Resize(GetLastUsedRow(wksParent).Row - rngDestination.Row + 1, Recordset.Fields.Count).ClearContents
                    End If
                End If
            Case ClearDestinationType.ClearAll
                'Clear used range, leaving formatting.
                wksParent.UsedRange.ClearContents
            Case ClearDestinationType.ClearFirstRow
                Set rng = GetUsedRowByStartCell(rngDestination)
                If Not rng Is Nothing Then
                    rng.ClearContents
                End If
            Case ClearDestinationType.DeleteAll
                strTemp = rngDestination.Address
                wksParent.UsedRange.EntireRow.Delete
                wksParent.UsedRange.EntireColumn.Delete
                Set rngDestination = wksParent.Range(strTemp)
            Case ClearDestinationType.DeleteUsedRangeBelowHeader
                If Not IsEmpty(GetLastUsedRow(wksParent)) And Recordset.Fields.Count > 0 Then
                    strTemp = rngDestination.Address
                    If Not GetLastUsedRow(wksParent) Is Nothing And Not GetLastUsedColumn(wksParent) Is Nothing Then
                        If rngDestination.Row <= GetLastUsedRow(wksParent).Row And rngDestination.Column <= GetLastUsedColumn(wksParent).Column Then
                            rngDestination.Resize(GetLastUsedRow(wksParent).Row - rngDestination.Row + 1, GetLastUsedColumn(wksParent).Column - rngDestination.Column + 1).Delete (xlUp)
                        End If
                    End If
                    Set rngDestination = wksParent.Range(strTemp)
                End If
            Case ClearDestinationType.NoAction
        End Select
        
        lRetVal = rngDestination.CopyFromRecordset(Recordset, MaxRows, MaxColumns) 'Returns long
        Set CopyRecordsetToExcelRange = rngDestination.Resize(lRetVal, Recordset.Fields.Count)
'        Call rngDestination.CopyFromRecordset(Recordset, MaxRows, MaxColumns) 'Returns long
            
        If AutoFit = True Then
            wksParent.UsedRange.EntireColumn.AutoFit
            wksParent.UsedRange.EntireRow.AutoFit
        End If
        
        If AutoFormat = True Then
            wksParent.UsedRange.NumberFormat = "General"
        End If
        
        If ResetCursor = True Then
            Recordset.MoveFirst 'Can cause requery 'Recordset.Requery 'May not be able to move cursor if Curortype is adOpenForwardOnly.
        End If
    End If
    
errsub:
    m_Application.EnableEvents = bEnableEvents
    m_Application.ScreenUpdating = bScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Public Function CopyRecordsetArrayToExcelRange(ByVal vRecordSet As Variant, rngDestination As Object, ClearDestination As ClearDestinationType, Optional AutoFit As Boolean = True, Optional AutoFormat As Boolean = False, Optional Transpose As Boolean = False) As Object 'Excel.Range
'Copies recordset to Excel sheet given recorset and start cell of range destination. Derived from Excel help for CopyFromRecordset.
'Returns modified range
'Possibly also use "SELECT * INTO" syntax to accomplish the same. 'http://support.microsoft.com/kb/200427
'In Excel can use transpose on array before passing in to reorient:
'shtTest.Cells(1, 1).Resize(UBound(vRecordset, 2) + 1, UBound(vRecordset, 1) + 1) = Application.WorksheetFunction.Transpose(vRecordset)
    On Error GoTo errsub

    Dim bScreenUpdating As Boolean
    Dim bEnableEvents As Boolean
    Dim rngAffected As Object 'Excel.Range
    Dim rng As Object 'Excel.Range
    Dim wksParent As Object 'Excel.Worksheet
    Dim strTemp As String
    
    bScreenUpdating = m_Application.ScreenUpdating
    bEnableEvents = m_Application.EnableEvents
    Set wksParent = rngDestination.Parent
    Set rngDestination = rngDestination(1)  'Only use first cell.
    
    If IsEmpty(vRecordSet) = False Then
        If m_Application.ScreenUpdating = True Then
            m_Application.ScreenUpdating = False 'Causes "Filling Cells" status message in Excel 2007.
        End If
        m_Application.EnableEvents = False 'Causes application flicker if on.
        
        Select Case ClearDestination
            Case ClearDestinationType.ClearUsedRangeBelowHeader
                'Clear Range, leaving formatting.
                If wksParent.UsedRange.Rows.Count > 0 And wksParent.UsedRange.Columns.Count > 0 Then
                    Call rngDestination.Resize(wksParent.UsedRange.Rows.Count, wksParent.UsedRange.Columns.Count).ClearContents '.Delete(xlShiftUp) using delete might help to keep used range small.
                End If
'                wksParent.UsedRange.ClearContents
            Case ClearDestinationType.ClearColumnsBelowHeader
                'Clear columns in destination area, leaving formatting.
                If wksParent.UsedRange.Rows.Count > 0 And UBound(vRecordSet, 2) + 1 > 0 Then
                    Call rngDestination.Resize(wksParent.UsedRange.Rows.Count, UBound(vRecordSet, 1) + 1).ClearContents '.Delete(xlShiftUp) using delete might help to keep used range small.
                End If
            Case ClearDestinationType.ClearAll
                'Clear used range, leaving formatting.
                wksParent.UsedRange.ClearContents
            Case ClearDestinationType.ClearFirstRow
                Set rng = GetUsedRowByStartCell(rngDestination)
                If Not rng Is Nothing Then
                    rng.ClearContents
                End If
            Case ClearDestinationType.DeleteAll
                strTemp = rngDestination.Address
                wksParent.UsedRange.EntireRow.Delete
                wksParent.UsedRange.EntireColumn.Delete
                Set rngDestination = wksParent.Range(strTemp)
            Case ClearDestinationType.DeleteUsedRangeBelowHeader
                If wksParent.UsedRange.Rows.Count > 1 And wksParent.UsedRange.Columns.Count > 0 Then
                    strTemp = rngDestination.Address
                    wksParent.UsedRange.Resize(wksParent.UsedRange.Rows.Count - 1).Offset(1).EntireRow.Delete
                    Set rngDestination = wksParent.Range(strTemp)
                End If
            Case ClearDestinationType.NoAction
        End Select
        
        'Transpose required to maintain expected functionality.  Code is opposite of what might be expected.
        If Transpose = False Then
            vRecordSet = TransposeArray(vRecordSet)
'            vRecordSet = m_Application.Transpose(vRecordSet) 'Leaves #N/A in bounds of array.
        End If
        Set rngAffected = rngDestination.Resize(UBound(vRecordSet, 1) + 1, UBound(vRecordSet, 2) + 1)
        rngAffected = vRecordSet
            
        If AutoFit = True Then
            wksParent.UsedRange.EntireColumn.AutoFit
            wksParent.UsedRange.EntireRow.AutoFit
        End If
        
        If AutoFormat = True Then
            wksParent.UsedRange.NumberFormat = "General"
        End If
        
        Set CopyRecordsetArrayToExcelRange = rngAffected
    End If
    
errsub:
    m_Application.EnableEvents = bEnableEvents
    m_Application.ScreenUpdating = bScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Public Function CreateDBFile(strPath As String) As Object
'Creates empty DB file and passes back reference to created DB.
'http://support.microsoft.com/kb/150418
'Note: MUST set returned object to nothing to close DB after modifaction is complete before performing operations on DB.  Set objADOXDB = Nothing  'close DB
'Note: Pre-Creation of a CSV may not be required if INSERT INTO ISAM syntax is used.

    On Error GoTo errsub
    
    Dim objADOXDatabase As Object
    Set objADOXDatabase = CreateObject("ADOX.Catalog")

    If IsFileOpen(strPath) Then
        Call MsgBox("Close the following file to proceed." & vbCrLf & vbCrLf & strPath, vbInformation + vbOKOnly, "Error")
        GoTo errsub
    End If
                            
    If DoesFilePathExist(strPath) = True Then
        Call Kill(strPath)
    End If
    'Could also check that path folders exists...
    Select Case GetDBDataFileTypeByExtension(strPath)
        Case eDataFileProvider.CSVFile, eDataFileProvider.MSExcel2003XLS, eDataFileProvider.MSExcel2007XLSX, eDataFileProvider.MSExcel2007XLSM, eDataFileProvider.MSExcel2007XLSB, eDataFileProvider.MSAccess2003, eDataFileProvider.MSAccess2007, eDataFileProvider.MSProject2003SP3
            Call objADOXDatabase.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath)
'        Case eDataFileProvider.MSAccess2000 'Access 2000 4.0 DB format = 5 'Lists Engine Type which can be determined with "connection.Properties("Jet OLEDB:Engine Type")"
'            Call objADOXDatabase.Create("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strPath & ";Jet OLEDB:Engine Type=5;")
        Case Else
            Debug.Assert False 'Add more?
    End Select
    
    Set CreateDBFile = objADOXDatabase
    Exit Function
    
errsub:
    Set objADOXDatabase = Nothing
    If Err.Number <> 0 Then
        Err.Raise Err.Number
    End If
End Function

Private Function BulkInsert(TableName As String, strFilePath As String)
'http://sqlserver2000.databases.aspfaq.com/how-do-i-load-text-or-csv-file-data-into-sql-server.html
'http://msdn.microsoft.com/en-us/library/ms188365.aspx
'May not be available for MS Access, instead export to CSV, then INSERT/UPDATE from CSV.
    
    Debug.Assert False 'Untested.
    Dim strCommandText As String
    
    'BULK INSERT YOURTABLENAME FROM 'yourfilename' WITH (FIELDTERMINATOR = '|',TABLOCK)
    strCommandText = "BULK INSERT" & TableName & "FROM" & strFilePath & "WITH (FIELDTERMINATOR = '|'"
    'INSERT INTO [AccessTable] SELECT * FROM [MS Access;DATABASE=D:\My Documents\db2.mdb].[Table2]
    'INSERT INTO [AccessTable] SELECT * FROM [MS Access;DATABASE=D:\My Documents\db2.mdb;PWD=password].[Table2]
End Function

Public Function CreateTableFromObjects(objADOXDBSource As Object, oRecordSet As Object, strTableName As String) As Boolean
'Given recordset, builds table (named strTableName) based on recordset field names.  Adds table to connected ADOXDB.
'Use with CreateTableFromObjects() or just use SelectIntoDBFromFile() to perform equivalent of bulk insert operation that creates and populates table.

    On Error GoTo errsub
    
    Dim Field As Object
    Dim objADOXTable As Object 'ADOX.Table
    
    Set objADOXTable = CreateObject("ADOX.Table")
    
    With objADOXTable
        objADOXTable.name = "[" & strTableName & "]"
        For Each Field In oRecordSet.Fields
            Call objADOXTable.Columns.Append(Field.name, Field.Type, Field.DefinedSize)
        Next
    End With
    
    Call objADOXDBSource.Tables.Append(objADOXTable)
    CreateTableFromObjects = True
    
errsub:
    Set objADOXTable = Nothing
    Set objADOXDBSource = Nothing 'Closes it.
End Function

Public Function ExportRecordSetToXMLFile(oRecordSet As Object, strFilePath As String) As String
'Untested
Debug.Assert False
'Write out oRecordSet to strTempFile in CSV format. Returns file path
'http://bytes.com/topic/visual-basic/answers/775966-how-update-records
    On Error GoTo errsub
    
'    Dim dom_document As Object 'DOMDocument

    If DoesFilePathExist(strFilePath) = True Then
        Call Kill(strFilePath)
    End If
    
    If Not oRecordSet Is Nothing Then
        Call oRecordSet.Save(strFilePath, PersistFormatEnum.adPersistXML) 'Save file to XML using Excel.
'        Call oRecordSet.Save(dom_document, PersistFormatEnum.adPersistXML)
'        Set dom_document = CreateObject("MSXML.DOMDocument") 'New DOMDocument
        ExportRecordSetToXMLFile = strFilePath
    End If
    
errsub:
    If Err.Number <> 0 Then
'        If DoesFilePathExist(strFilePath) = true Then 'If error delete file
'            Call Kill(strFilePath)
'        End If
        Err.Raise Err.Number
    End If
End Function

Public Function ExportRecordSetToCSVFile(oRecordSet As Object, strFilePath As String) As String
'Write out oRecordSet to strTempFile in CSV format.  Returns file path.

    On Error GoTo errsub
    
    Dim objFSO As Object
    Dim objFile As Object 'TextSteam
    Dim Field As Object
    Dim strTemp As String

    If DoesFilePathExist(strFilePath) = True Then
        Call Kill(strFilePath)
    End If
    
    Set objFSO = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
    Set objFile = objFSO.CreateTextFile(strFilePath)
    
    If Not oRecordSet Is Nothing Then
        For Each Field In oRecordSet.Fields
            strTemp = strTemp & Field.name & ", "
        Next
        strTemp = Left(strTemp, Len(strTemp) - 2) 'Remove trailing comma

        Call objFile.Writeline(strTemp) 'Write Header
        Call objFile.Writeline(oRecordSet.GetString(, , ", ", vbCrLf, vbNullString))  'Write Data
        ExportRecordSetToCSVFile = strFilePath
    End If
    
errsub:
    Set Field = Nothing
    Set objFile = Nothing
    Set objFSO = Nothing
    
    If Err.Number <> 0 Then
        If DoesFilePathExist(strFilePath) = True Then 'If error delete file
            Call Kill(strFilePath)
        End If
        Err.Raise Err.Number
    End If
End Function

Private Function WriteCSVSchemaIni(CSVFilePath As String) As Boolean
'Schema.ini required in the same location as CSV in order to force correct data type import.
'Created and used because CSV driver doesn't always import the correct data types.
'Simple method imports all data as text when columns are not explicitly defined.
'Can also specify fields with data types: Short, Long, Currency, Single, Double, DateTime, Memo
'http://msdn.microsoft.com/en-us/library/ms709353(v=vs.85).aspx
'http://msdn.microsoft.com/en-us/library/ms974559
'Format: Col(n)=<column name> <data type> <Width width>
    On Error GoTo errsub
    Dim strFileName As String
    Dim strFilePath As String
    
    Call DeleteTempFile(SchemaINIPath)
    
    strFilePath = ParsePath(CSVFilePath, Path) & "Schema.ini"
    strFileName = ParsePath(CSVFilePath, FileName)
    
    If WritePrivateProfileString(strFilePath, strFileName, "ColNameHeader", "True") Then
        If WritePrivateProfileString(strFilePath, strFileName, "Format", "CSVDelimited") Then
            If WritePrivateProfileString(strFilePath, strFileName, "MaxScanRows", "0") Then '0 - look at the entire file and set type based on majority.  This can set the wrong value if data is mixed: http://support.microsoft.com/kb/282263
        '    Call WritePrivateProfileString(CSVFilePath, strFileName, "DateTimeFormat", "dd-MMM-yyyy")
        '    Call WritePrivateProfileString(CSVFilePath, strFileName, "Col1", "A DateTime")
        '    Call WritePrivateProfileString(CSVFilePath, strFileName, "Col2", "B Text Width 100")
                SchemaINIPath = strFilePath
                WriteCSVSchemaIni = True
            End If
        End If
    End If
    Exit Function
    
errsub:
    'Read only file location?
End Function

Private Function WritePrivateProfileString(ByVal strFileName As String, ByVal strSection As String, ByVal strKey As String, ByVal strValue As String) As Boolean
'Write string to INI file.
    On Error GoTo errsub
    
    Dim lRetVal As Long
    
    lRetVal = WritePrivateProfileStringA(strSection, strKey, strValue, strFileName)
    
    If lRetVal > 0 Then
        WritePrivateProfileString = True
    End If
errsub:
End Function

Public Function UpdateDisconnectedRecordSet(oRecordsetSource As Object, DBTable As String) As Object 'Recordset
'Update a disconnected recordset then apply changes.  For use when data is in an array and not in a DB source.
'http://www.accessmonster.com/Uwe/Forum.aspx/databases-ms-access/43818/MS-Access-Mass-Bulk-Insert-into-a-table
'http://bytes.com/topic/access/answers/558797-insert-records-into-local-access-table-external-db
'See also the following link for three examples of how this is performed:
'http://support.microsoft.com/kb/184397

Debug.Assert False 'Untested stub, needs more work
''It is possible to get a recordset, disconnect, update/modify the recordset then apply the changes to the associated table.
''Test GBH
''Modified from:
''Save Recordset to XML file then reload in new custom recordset connected to active connection data source and update batch.
''http://bytes.com/topic/visual-basic/answers/775966-how-update-records
'Call oRecordsetSource.Save(strPath, PersistFormatEnum.adPersistXML)
'
'
'Other links for disconnected RecordSet()
'http://www.codeproject.com/KB/database/aadoclass.aspx
'
'
'    On Error GoTo errsub
'
'    Dim objDB As clsDBConnectivity
'    Dim strPath As String
'    Dim oRecordset As Object 'ADODB.Recordset
'    Set oRecordset = CreateObject("ADODB.recordset")
'
'    strPath = (VBA.Environ("TEMP") & "\SRNSTempCSVExport.xml")
'
'    If DoesFilePathExist(strPath) = True Then
'        Call Kill(strPath)
'    End If
'
'    Call oRecordsetSource.Save(strPath, PersistFormatEnum.adPersistXML)
'
'    With oRecordset 'Custom Recordset
'        .CursorLocation = CursorLocationEnum.adUseClient 'Need in order to be able to disconnect
'        .ActiveConnection = DBConnection()
'        .CursorType = CursorTypeEnum.adOpenStatic 'Needed to be able to update.
'        .LockType = LockTypeEnum.adLockBatchOptimistic
'        'Don't put in Options for improved performance.
'        .ActiveConnection = Nothing 'Disconnect
'        Call .Open(strPath) 'Works with default Recordset settings.
'        .ActiveConnection = DBConnection()
'        'Can I use query to another data source here as in?:
'        'SELECT * INTO [Data] FROM [Excel 8.0;DATABASE=E:\My Documents\Test.xls;HDR=No;IMEX=1].[Sheet1$]
'        .UpdateBatch    'Use instead of Batch()
'        .Close
'    End With
'
'errsub:
'    'Untested.
'
'    If Not oRecordset Is Nothing Then
'        If oRecordset.State = ObjectStateEnum.adStateOpen Then
'            oRecordset.Close
'        End If
'        Set oRecordset = Nothing
'    End If


'  ' Open Connection to Access Database
'   Dim conn As ADODB.Connection
'   Set conn = New ADODB.Connection
'   Dim connString As String
'   connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DBName & ";"
'   conn.Open connString
'
'   ' Open New Recordset based on Access Database Table
'   Dim rst As ADODB.Recordset
'   Set rst = New ADODB.Recordset
'   With rst
'       .ActiveConnection = conn
'       .CursorLocation = adUseClient
'       .CursorType = adOpenStatic
'       .LockType = adLockBatchOptimistic
'       .Open DBTable
'   End With
'
'   ' Disconnect Recordset
'   rst.ActiveConnection = Nothing
'   conn.Close
'
' [ …Code to parse data and populate Recordset]
'
'  ' Re-Open Connection & Upload Recordset to Access Table
'   conn.Open connString
'   rst.ActiveConnection = conn
'   rst.UpdateBatch 'Use instead of Batch()
'   rst.Close

End Function

Public Function UpdateExcelRangeText(strRange As String, strValue As String) As Boolean
    'In ConnectToDB(), must have mode set to "UpdateMode" .
    'Can update multiple columns by expanding range and adding more F specifiers F1, F2, F3...
    'strCommandText = "UPDATE [Sheet1$A2:A2] SET F1='11'"
    Dim strCommandText As String
    strCommandText = "UPDATE [" & strRange & "] SET F1='" & strValue & "'"
    Call ExecuteCommand(strCommandText)
    UpdateExcelRangeText = True
End Function

Public Function GetArrayFromRecordSet(oRecordSet) As Variant()
    GetArrayFromRecordSet = oRecordSet.GetRows
End Function

Public Function GetRecordSetValue(Value As Variant) As Variant
'GetRecordSetValue(RecordSet("<value>")) to retuen Empty if Null
    GetRecordSetValue = IIf(VBA.IsNull(Value), Empty, Value)
End Function

Public Function SelectTop1RecordSet(Optional TableName As String) As Object
    'If no passed in table uses table that was connected to in the connection string.
    Dim strCommandText As String
    
    If TableName = vbNullString Then
        TableName = DBQueryTableName
    End If
    
    Select Case DBDataSourceType
        Case eDataProvider.Oracle, eDataProvider.MSOracle
            strCommandText = "SELECT * FROM(SELECT * FROM " & TableName & " DataTable)) WHERE ROWNUM <= 1"
        Case Else 'T-SQL
            strCommandText = "SELECT TOP 1 * FROM " & TableName & " DataTable"
    End Select
    
    Set SelectTop1RecordSet = ExecuteRecordset(strCommandText) 'Recordset
'    SelectTop1Record = ExecuteArray(strCommandText) 'Array
End Function

Public Function SelectTop1Array(Optional TableName As String) As Variant
    'If no passed in table uses table that was connected to in the connection string.
    Dim strCommandText As String
    
    If TableName = vbNullString Then
        TableName = DBQueryTableName
    End If
    
    Select Case DBDataSourceType
        Case eDataProvider.Oracle, eDataProvider.MSOracle
            strCommandText = "SELECT * FROM(SELECT * FROM " & TableName & ")) WHERE ROWNUM <= 1"
        Case Else 'T-SQL
            strCommandText = "SELECT TOP 1 * FROM [" & TableName & "]"
    End Select
    
    SelectTop1Array = ExecuteArray(strCommandText) 'Array
End Function

Public Function SelectAllRecordsRecordset(Optional TableName As String, Optional bStoreRecordCount As Boolean = False) As Object 'ADODB.Recordset 'strExcelSheetName As String, Optional bStoreRecordCount As Boolean = False) As Object 'ADODB.Recordset
'http://support.microsoft.com/kb/257819
    Dim strCommandText As String
    
    If TableName = vbNullString Then
        TableName = DBQueryTableName
    End If
    
    strCommandText = "SELECT * FROM " & TableName
    Set SelectAllRecordsRecordset = ExecuteRecordset(strCommandText, bStoreRecordCount)
End Function

Public Function SelectAllRecordsArray(Optional TableName As String, Optional bStoreRecordCount As Boolean = False) As Variant
'http://support.microsoft.com/kb/257819
    Dim strCommandText As String
    
    If TableName = vbNullString Then
        TableName = DBQueryTableName
    End If
    
    strCommandText = "SELECT * FROM " & TableName
    SelectAllRecordsArray = ExecuteArray(strCommandText, bStoreRecordCount)
End Function

Public Function CopyWorksheet(wksSource As Object, wksDestination As Object, Optional IncludeHeader As Boolean = False, Optional AutoFit As Boolean = True, Optional AutoFormat As Boolean = False) As Object 'Excel.Range
'Public Function CopyWorksheet(wksSource As Excel.Worksheet, wksDestination As Excel.Worksheet, Optional IncludeHeader As Boolean = False, Optional AutoFit As Boolean = True, Optional AutoFormat As Boolean = False) As Excel.Range
'Copies wksSource to wksDestination returning affected data range.
    Dim vRecordSet As Variant

    vRecordSet = SelectAllRecordsArray("[" & wksSource.name & "$]")
    If IsEmpty(vRecordSet) = False Then
        If IncludeHeader = True Then
            Call CopyRecordsetHeaderToExcelRange(wksDestination.Range("A1"), True, AutoFit, AutoFormat)
        End If
        Set CopyWorksheet = CopyRecordsetArrayToExcelRange(vRecordSet, wksDestination.Range("A2"), ClearUsedRangeBelowHeader, AutoFit)
    End If
End Function

Public Function SelectAllRecordsByExcelRangeRecordset(rngData As Object, Optional bStoreRecordCount As Boolean = False) As Object
'Pass in range and query will return recordset.
'http://support.microsoft.com/kb/257819
'[<Protocol Name>$<Range>]
'Class doesn't support using this to return just one row as of 5/28/10
    Dim strCommandText As String

    strCommandText = "SELECT * FROM " & GetRangeQueryAddress(rngData)
    
'    Set SelectAllRecordsByRange = ExecuteRecordset(strCommandText, bStoreRecordCount)
    Set SelectAllRecordsByExcelRangeRecordset = ExecuteRecordset(strCommandText)
End Function

Public Function GetSelectQueryStringFromExcelRangeRecordset(ByVal rngSourceHeader As Object, Optional Alias As String = "") As String
'Pass in range header containing fields to return from query.
    Dim strCommandText As String
    Dim rng As Object 'Excel.Range
    
    strCommandText = "SELECT "
    
    For Each rng In rngSourceHeader
        If Alias = vbNullString Then
            strCommandText = strCommandText & "[" & rng.Value & "], "
        Else
            strCommandText = strCommandText & Alias & ".[" & rng.Value & "], "
        End If
    Next
    
    strCommandText = Left(strCommandText, Len(strCommandText) - 2) 'Remove trailing comma
    
    If Alias = vbNullString Then
        strCommandText = strCommandText & " FROM " & DBQueryTableName
    Else
        strCommandText = strCommandText & " FROM " & DBQueryTableName & VBA.Space(1) & Alias
    End If
    
    GetSelectQueryStringFromExcelRangeRecordset = strCommandText
End Function

Public Function SelectRecordsByExcelRangeRecordset(ByVal rngSourceHeader As Object) As Object
'Pass in range header containing fields to return from query.
    Dim strCommandText As String
    
    strCommandText = GetSelectQueryStringFromExcelRangeRecordset(rngSourceHeader)
    
'    Debug.Print strCommandText
    Set SelectRecordsByExcelRangeRecordset = ExecuteRecordset(strCommandText)
End Function
            
Public Function SelectAllRecordsByExcelRangeArray(rngData As Object, Optional bStoreRecordCount As Boolean = False) As Variant
'Pass in range and query will return recordset.
'http://support.microsoft.com/kb/257819
'[<Protocol Name>$<Range>]
'Warning may not work in Excel 2007.  Work around is to dynamically create named range to query then remove when finished.
    Dim strCommandText As String

    strCommandText = "SELECT * FROM " & GetRangeQueryAddress(rngData)
    
'    Set SelectAllRecordsByRange = ExecuteRecordset(strCommandText, bStoreRecordCount)
    SelectAllRecordsByExcelRangeArray = ExecuteArray(strCommandText)
End Function

Public Function GetRangeQueryAddress(rngData As Excel.Range) As String
    If Not rngData Is Nothing Then
        GetRangeQueryAddress = "[" & rngData.Parent.name & "$" & rngData.Address(False, False) & "]"
    End If
End Function

Public Function DropTable(strTable As String) As Boolean
    'Removes table from database
    On Error Resume Next 'Can error if table doesn't exist.
    DropTable = ExecuteCommand("DROP TABLE " & strTable)
End Function

Public Function DeleteTable(strTable As String) As Boolean
    'Removes all rows from a table
    On Error Resume Next 'Can error if table doesn't exist.
    DeleteTable = ExecuteCommand("DELETE " & strTable)
'    DeleteTable = Execute("DELETE FROM " & strTable)
'    DropTable = Execute("DELETE * FROM " & strTable)
End Function

Public Function CopyTable(strOriginalTable As String, strNewTable As String) As Boolean
'Copy one table into another new table - usually for testing
'http://www.experts-exchange.com/Databases/Microsoft_SQL_Server/Q_20412715.html
    
'    Debug.Print "SELECT * INTO " & strNewTable & " FROM " & strOriginalTable
    CopyTable = ExecuteCommand("SELECT * INTO " & strNewTable & " FROM " & strOriginalTable)
    
'    select *
'    into <other table name>
'    from <tablename>
'    where ...
'
'    to achieve this you have to make sure that th option 'select into/bulkinsert is enabled.  Which you can do like this:
'    exec sp_dboption <databasename>, 'select into/bulkcopy', true
'
'    SELECT * INTO [ttt] FROM [xxx] WHERE id is null
'    --this is a type of BulkCopy
'
'    if  this gives an error then probably
'    your bulcopy option is set off
'
'    EXEC sp_dboption 'dbname', 'isbulkcopy', 'true'
End Function

Private Function DetachDB(DBName As String) As Boolean
'Use DetachDBSYS_SP()
'Detatch DB file *.MDF - requires admin rights
    On Error GoTo errsub
    Dim strCommand As String
    
    strCommand = "DBCC DETACHDB('" & DBName & "')"

    Call ExecuteCommand(strCommand)
    DetachDB = True
errsub:
End Function

Private Function AttachSingleFileDB(DBName As String, DBFilePath As String) As Boolean
'Use AttachSingleFileDBSYS_SP()
'Attach DB - Calls system stored procedure sp_attach_single_file_db to attach *.MDF DB file
'Requires admin rights
On Error GoTo errsub
    Dim strCommand As String
    
    If DoesFilePathExist(DBFilePath) = True And DBName <> vbNullString Then
        strCommand = "CREATE DATABASE [" & DBName & "] " & _
        "ON (FILENAME ='" & DBFilePath & "') " & _
        "FOR ATTACH"
    
        Call ExecuteCommand(strCommand)
        AttachSingleFileDB = True
    End If
errsub:
End Function

Public Function SelectIntoDBFromFile(strDataSource As String, strDataTable As String, strPathtoTextFileSource As String) As Boolean
    'Creates new DB and inserts records from txt or csv.
    'ADO guesses on fields types based on data, which can be incorrect for fields such as dates.
    'Previous table is deleted.
    'http://support.microsoft.com/kb/262537
    'http://social.msdn.microsoft.com/Forums/en-US/adodotnetdataproviders/thread/c7cbc0e5-e7f6-44c5-a382-1595034faa44
    
'    On Error GoTo errsub
'
'    Dim objDB As clsDBConnectivity
    Dim strCommand As String
'
'    Set objDB = New clsDBConnectivity
'
'    If objDB.ConnectToDB(FileBased, strDataSource) = True Then
        Call DropTable(strDataTable)
        
        'http://www.w3schools.com/sql/sql_select_into.asp
        strCommand = "SELECT * INTO [" & strDataTable & "] FROM " & "[Text;Database=" & ParsePath(strPathtoTextFileSource, Path) & ";HDR=YES].[" & ParsePath(strPathtoTextFileSource, FileName) & "]"
        
        Call ExecuteCommand(strCommand)
        SelectIntoDBFromFile = True
'    End If

errsub:
'    Set objDB = Nothing
End Function

Private Function SelectIntoConnectedDB()
Debug.Assert False 'Untested work in progress.
Exit Function
''http://support.microsoft.com/kb/321686
''http://support.microsoft.com/kb/200427
''http://www.mikesdotnetting.com/Article/79/Import-Data-From-Excel-to-Access-with-ASP.NET
''Can also switch the order so that the dynamic connection is made to the INTO DB.
'    On Error GoTo errsub
'
'    Dim objDB As clsDBConnectivity
'    Set objDB = New clsDBConnectivity
'    Dim strDBPath As String
'    Dim strDataTableDestination As String
'    Dim strDataTable As String
'    Dim vRecordSet As Variant
'    Dim strCommand As String
'    Dim strFilePath As String
'
'    strDataTableDestination = "TestFromExcel2"
'    strDBPath = ThisWorkbook.Path & "\PMInputs.mdb"
'    strFilePath = ThisWorkbook.FullName
'    strDataTable = "FromExcel"
'
'    If Not objDB.CreateDBFile(strDBPath) Is Nothing Then
'        'Need to delete table if it already exists using ADOX.
'
'        If objDB.ConnectToDB(FileBased, strDBPath) = True Then
'
'    '        Call objDB.DropTable(strDataTableDestination)
'
'            Select Case UCase(ParsePath(strFilePath, FileExtension))
'    '        strCommand = "SELECT * INTO [MS Access;Database=" & Access & "].[New Table] FROM [Sheet1$]"
'    '        strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & "[MS Access;Database=" & ParsePath(PathtoTextFile, Path) & ";HDR=YES].[" & ParsePath(PathtoTextFile, FileName) & "]"
'                Case "CSV" 'Working
'                    strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & GetISAMName(FileBased, strFilePath, , strDataTable)
'                Case "XLS"
'                    '[<Full path to Microsoft Access database>].[<Table Name>]
'                    '[ODBC;<ODBC Connection String>].[<Table Name>]
'                    '[<ISAM Name>;<ISAM Connection String>].[<Table Name>]
'    '                ISAM = Indexed Sequential Access Method
'                    '[Text;Database=C:\\Path\;HDR=YES].[txt.csv]
'                    '[Excel 8.0;Database=c:\\customers.xls;HDR=Yes].[Sheet1]
'                    '[MS Access;Database=c:\\customers.mdb].[New Table]
'                    '[ODBC;Driver=SQL Server;SERVER=XXX;DATABASE=Pubs;UID=<username>;PWD=<strong password>;].[RemoteShippers]
'
'    '                strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & "[ODBC;Driver={Microsoft Excel Driver (*.xls)}; DriverId=790; Dbq=" & strFilePath & ";].[" & strDataTable & "$] "
'
'    '                strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & "[Excel 8.0;HDR=Yes;IMEX=1;MaxScanRows=16;Database=" & strFilePath & "].[" & strDataTable & "$]"
'                    'Working:
'                    strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & GetISAMName(FileBased, strFilePath, , strDataTable)
'                    '.Properties("Extended Properties") = "Excel 8.0;HDR=" & IIf(bFileHasHeader, "Yes", "No") & ";IMEX=1;MaxScanRows=16" 'MaxScanRows=8-16
'                Case "MDB" 'Should be working:
'                    strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & GetISAMName(FileBased, strFilePath, , strDataTable)
'                Case "XLSM"
'                    'Not working.  Need ISAM driver name for Excel 12 if there is one.
'                    'strCommand = "SELECT * INTO [" & strDataTableDestination & "] FROM " & "[Excel 12.0 Macro;HDR=Yes;IMEX=1;MaxScanRows=16;Database=" & strFilePath & "].[" & strDataTable & "$]"
'            End Select
'    '        Debug.Print strCommand
'
'            Call objDB.ExecuteRecordset(strCommand)
'        End If
'    End If
'
'errsub:
'    Debug.Assert False
'    Set objDB = Nothing
End Function

Private Function SelectIntoTableFromExternalWorkbook()
    Debug.Assert False 'Untested
'Creates new DB and inserts records.  Table must not exist first.
'http://support.microsoft.com/kb/316934
'SELECT * INTO [Excel 8.0;Database=C:\Book1.xls].[Sheet1] FROM [MyTable]
End Function

Public Function InsertIntoDBFromFile(strDataSource As String, strDataTable As String, strPathtoTextFileSource As String) As Boolean
'Appends records to existing DB from txt or csv.
'Use with CreateTableFromObjects() or just use SelectIntoDBFromFile() to perform equivalent of bulk insert operation that creates and populates table.
    'http://support.microsoft.com/kb/262537
    'http://social.msdn.microsoft.com/Forums/en-US/adodotnetdataproviders/thread/c7cbc0e5-e7f6-44c5-a382-1595034faa44
    On Error GoTo errsub
    
    Dim objDB As clsDBConnectivity
    Dim strCommand As String
    
    Set objDB = New clsDBConnectivity
    
    If objDB.ConnectToDB(FileBased, strDataSource) = True Then
        'Insert into table1 the contents of textfile.txt
        strCommand = "INSERT INTO [" & strDataTable & "] SELECT * FROM " & "[Text;Database=" & ParsePath(strPathtoTextFileSource, Path) & ";HDR=YES].[" & ParsePath(strPathtoTextFileSource, FileName) & "]"
        Call objDB.ExecuteCommand(strCommand)
        InsertIntoDBFromFile = True
    End If

errsub:
    Set objDB = Nothing
End Function

Private Function InsertTableFromExternalWorkbook()
    Debug.Assert False 'Untested
'http://support.microsoft.com/kb/316934
'INSERT INTO [Sheet1$] IN 'C:\Book1.xls' 'Excel 8.0;' SELECT * FROM MyTable"
End Function

Private Function SelectByExcelNamedRange(rngSource As Object, strCommandText As String) As Object
    Debug.Assert False 'Untested
'   Call SelectAllRecordsByExcelRangeRecordset
'    On Error GoTo errsub
'    Dim strDataTable As String
'
'    strDataTable = "ProModelTempNamedRange"
'    Call rngSource.Parent.Parent.Names.Add(strDataTable, rngSource) 'Will reset a previous value if already exists.
'    strCommandText = GetQueryUniqueAreas(strDataTable)
'    Set SelectByExcelRange = objDB.ExecuteRecordset(strCommandText)
    
'errsub: 'Fall through intentional
'    Call rngSource.Parent.Parent.Names(strDataTable).Delete

End Function

Public Function GetValidDataTableName(Optional bAllowValidDataTableSelection As Boolean = False) As Boolean
'http://support.microsoft.com/kb/186246
'http://support.microsoft.com/kb/257819
'Get sheets in Excel workbook without opening.  Does not list Hidden or VeryHidden worksheets.
'http://www.codeproject.com/KB/aspnet/getsheetnames.aspx
    Dim strResultOriginal As String
    Dim strResult As String
    Dim lCursor As Long
    Dim bScreenUpdating As Boolean
    Dim oRecordSet As Object 'ADODB.Recordset
    Dim i As Long
    
    Const xlDefault = -4143
    
    'Save default values
    If m_Application.name = "Microsoft Excel" Then
        lCursor = m_Application.Cursor
    End If
    If m_Application.name = "Microsoft Excel" Or m_Application.name = "Microsoft Visio" Then
        bScreenUpdating = m_Application.ScreenUpdating
    End If
    
    'Look for a particular table name in schema (should return only one record if found)
    'Array values to pass for OpenSchema(SchemaEnum.adSchemaTables): TABLE_CATALOG,TABLE_SCHEMA,TABLE_NAME,TABLE_TYPE,TABLE_GUID,DESCRIPTION,TABLE_PROPID,DATE_CREATED,DATE_MODIFIED
    If VBA.InStr(1, DBTable, VBA.Space(1)) <> 0 Then
        Set oRecordSet = m_DBConnection.OpenSchema(SchemaEnum.adSchemaTables, Array(Empty, DBSchema, "'" & DBTable & "'"))
    Else
        Set oRecordSet = m_DBConnection.OpenSchema(SchemaEnum.adSchemaTables, Array(Empty, DBSchema, DBTable))
    End If
    
    'If not found, query all tables and add to list for selection.
    If oRecordSet.EOF Then
        DBTable = Empty
        Me.cmbDataTableSelection.Clear
        Set oRecordSet = m_DBConnection.OpenSchema(SchemaEnum.adSchemaTables)
        Do While Not oRecordSet.EOF
            strResultOriginal = VBA.Trim(oRecordSet.Fields("Table_Name").Value)  'If oRecordSet("TABLE_TYPE").Value = "SYSTEM_TABLE" Then
'            strResult = Replace(strResultOriginal, "#", ".") 'Simbryo has these.
'            strResult = Replace(strResult, "'", vbNullString)
            If (DBDataFileIsExcel = True) And ((VBA.Right(strResultOriginal, 1) = "$") Or (Right(strResultOriginal, 2) = "$'")) Then
                'Workbook name can contain: apostrophe , so replace with character, remove others, then restore apostrophe
                strResultOriginal = VBA.Replace(strResultOriginal, "''", Chr(1))
                strResultOriginal = VBA.Replace(strResultOriginal, "'", vbNullString)
                strResultOriginal = VBA.Replace(strResultOriginal, Chr(1), "'")
                strResult = strResultOriginal
'                strResult = VBA.Replace(strResultOriginal, "''", Chr(1))
'                strResult = VBA.Replace(strResult, "'", vbNullString) 'Remove ''
'                strResult = VBA.Replace(strResult, Chr(1), "'")
                If VBA.Right(strResult, 1) = "$" Then
                    strResult = VBA.Left(strResult, VBA.Len(strResult) - 1) 'Worksheet name itself can contain $ sign, so only strip off last one.
                End If
            Else
                strResult = strResultOriginal
            End If
            
            'Add to ComboBox in sorted order.
            If UBound(Me.cmbDataTableSelection.List) > -1 Then 'Not first item
                For i = LBound(Me.cmbDataTableSelection.List) To UBound(Me.cmbDataTableSelection.List)
                    If UCase(cmbDataTableSelection.List(i, 0)) >= VBA.UCase(strResult) Then
                        Exit For
                    End If
                Next i
            End If
            Call Me.cmbDataTableSelection.AddItem(strResult, i) 'Normal Name
            Me.cmbDataTableSelection.List(i, 1) = strResultOriginal 'Real Name

            If bAllowValidDataTableSelection = False Then 'Retreive at least one valid default table to connect to.
                Exit Do
            End If
'            Debug.Print strResultOriginal
            oRecordSet.MoveNext
        Loop
        
        If Me.cmbDataTableSelection.ListCount > 1 Then 'Show if more than one to choose from.  Auto select one if only one.(there must be at least one in Excel workbook)
            If m_Application.name = "Microsoft Excel" Then
                m_Application.Cursor = xlDefault
            End If
            If m_Application.name = "Microsoft Excel" Or m_Application.name = "Microsoft Visio" And m_Application.ScreenUpdating = False Then
                m_Application.ScreenUpdating = True
            End If

            Me.Caption = "Database Connection"
            Me.frmDBSelection.Caption = "Available Tables:"
            Me.cmbDataTableSelection.ControlTipText = "Supplied table is invalid. Select table to open from the available list."
            Me.cmbDataTableSelection.ListIndex = 0 'set to first item in list.
            Me.Show vbModal 'Unload of form canceled with QueryClose so that we don't lose object references.
            DBTable = VBA.CStr(Me.cmbDataTableSelection.List(cmbDataTableSelection.ListIndex, 1)) 'Reset passed in value.
        Else
            DBTable = VBA.CStr(Me.cmbDataTableSelection.List(cmbDataTableSelection.ListCount - 1, 1)) 'Reset passed in value.
        End If
    End If
    
    GetValidDataTableName = True
    
errsub: 'Fall through intentional
    'Restore default values
    If m_Application.name = "Microsoft Excel" Then
        m_Application.Cursor = lCursor
    End If
    If m_Application.name = "Microsoft Excel" Or m_Application.name = "Microsoft Visio" Then
        m_Application.ScreenUpdating = bScreenUpdating
    End If
    
    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            oRecordSet.Close
        End If
        Set oRecordSet = Nothing
    End If
End Function

Public Function CopyRecordsetHeaderToExcelRange(rngDestination As Object, Optional ClearEntireHeaderRow As Boolean = False, Optional AutoFit As Boolean = True, Optional AutoFormat As Boolean = False) As Object
'Public Function CopyRecordsetHeaderToExcelRange(rngDestination As Object, Optional ClearEntireHeaderRow As Boolean = False, Optional AutoFit As Boolean = True, Optional AutoFormat As Boolean = False) As Excel.Range
'Returns range of results which can be set to range or variant array.
'Could add check to make sure array will fit using MaxRows/MaxColumns
'Some symbols do not not copy as expected such as: period(.)-> #
'    Dim MaxRows As Variant
'    Dim MaxColumns As Variant
    On Error GoTo errsub
    Dim bScreenUpdating As Boolean
    Dim bEnableEvents As Boolean
    Dim vHeaderArray As Variant
    
    bScreenUpdating = m_Application.ScreenUpdating
    bEnableEvents = m_Application.EnableEvents
'    MaxRows = (rngDestination.Parent.Rows.Count - rngDestination.row) + 1
'    MaxColumns = (rngDestination.Parent.Columns.Count - rngDestination.Column) + 1
    Set rngDestination = rngDestination(1)  'Only use first cell.
    
'    If Not Recordset Is Nothing Then
        If m_Application.ScreenUpdating = True Then
            m_Application.ScreenUpdating = False 'Causes "Filling Cells" status message in Excel 2007.
        End If
        m_Application.EnableEvents = False 'Causes application flicker if on.
        
        If ClearEntireHeaderRow = True Then
            rngDestination.EntireRow.ClearContents
        End If
        
        vHeaderArray = RecordSetFieldsArray(False)
        If IsEmpty(vHeaderArray) = False Then 'Can return nothing.
            Set CopyRecordsetHeaderToExcelRange = rngDestination.Resize(1, UBound(vHeaderArray) + 1)
            rngDestination.Resize(1, UBound(vHeaderArray) + 1) = vHeaderArray
            
            If AutoFit = True Then
                rngDestination.Parent.UsedRange.EntireColumn.AutoFit
                rngDestination.Parent.UsedRange.EntireRow.AutoFit
            End If
            
            If AutoFormat = True Then
                rngDestination.Parent.UsedRange.NumberFormat = "General"
            End If
        End If
'        CopyRecordsetHeaderToExcelRange = True
'    End If
    
errsub:
    m_Application.EnableEvents = bEnableEvents
    m_Application.ScreenUpdating = bScreenUpdating
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Private Function GetDBDataFileTypeByExtension(ByVal strPath As String) As eDataFileProvider
    strPath = UCase(ParsePath(strPath, FileExtension))
    Select Case strPath
        Case "MDB" 'Access 2000, 2003
            GetDBDataFileTypeByExtension = eDataFileProvider.MSAccess2003
        Case "ACCDB" 'Access 2007
            GetDBDataFileTypeByExtension = eDataFileProvider.MSAccess2007
        Case "XLSX"
                GetDBDataFileTypeByExtension = eDataFileProvider.MSExcel2007XLSX
        Case "XLSM"
                GetDBDataFileTypeByExtension = eDataFileProvider.MSExcel2007XLSM
        Case "XLSB"
                GetDBDataFileTypeByExtension = eDataFileProvider.MSExcel2007XLSB
        Case "XLS"
            'If newer host application, use newer driver to avoid file save delay required to work around memory leak with xls.
            'Doesn't work correctly on Vista with Excel 2007, so limited to Windows 7 or newer.
            If StrComp(m_Application.name, "Microsoft Excel") = False And m_Application.Version >= 12 And GetOSVersionNumber >= 6.1 Then  'Excel 2007 or later and Windows 7 or later.
                GetDBDataFileTypeByExtension = eDataFileProvider.MSExcel2007XLSM
            Else
                GetDBDataFileTypeByExtension = eDataFileProvider.MSExcel2003XLS
            End If
        Case "MPP"
            GetDBDataFileTypeByExtension = eDataFileProvider.MSProject2003SP3
        Case "CSV"
            GetDBDataFileTypeByExtension = eDataFileProvider.CSVFile
        Case "XML"
            GetDBDataFileTypeByExtension = eDataFileProvider.XMLFile
        Case Else 'Database
            GetDBDataFileTypeByExtension = eDataFileProvider.NA
    End Select
End Function

Private Function GetDataSourceCategory(ByVal DataProvider As eDataProvider) As eDataProviderCategory
    Select Case DataProvider
        Case eDataProvider.FileBased
            GetDataSourceCategory = eDataProviderCategory.Filebase
        Case eDataProvider.MSSQL2000, eDataProvider.MSSQLExpress2005, eDataProvider.MSSQLExpress2005Instance, eDataProvider.MSSQL2005, eDataProvider.MSSQL2008, eDataProvider.Composite45, eDataProvider.Composite50, eDataProvider.Oracle, eDataProvider.MSOracle, eDataProvider.PostgreSQL
            GetDataSourceCategory = eDataProviderCategory.database
        Case eDataProvider.DataSourceName
            GetDataSourceCategory = eDataProviderCategory.DataSourceName
    End Select
End Function

Private Function GetDataFileIsExcel(ByVal DBDataSourceType As eDataFileProvider) As Boolean
    Select Case DBDataSourceType
        Case eDataFileProvider.MSExcel2003XLS, eDataFileProvider.MSExcel2007XLSB, eDataFileProvider.MSExcel2007XLSM, eDataFileProvider.MSExcel2007XLSX
            GetDataFileIsExcel = True
    End Select
End Function

Private Function SetQueryTable()
    'A query table is a table in an Excel worksheet that's linked to an external data source, such as a SQL Server database, a Microsoft Access database, a Web page, or a text file.
    'To retrieve the most up-to-date data, the user can refresh the query table.
    'Creating Dynamic Reports with Query Tables in Excel
    'http://msdn.microsoft.com/en-us/library/aa188518(v=office.10).aspx
    'http://vbadud.blogspot.com/2007/12/query-table-with-excel-as-data-source.html
    'Performed some testing in attempt to integrate method with existing DB Connection code, though was unsuccessful.
    'LinkedIn post listed this method as faster than ADO connection using method CopyFromRecordset().
    'http://www.linkedin.com/groups/Does-anyone-know-how-pull-1871310.S.53228172?view=&gid=1871310&type=member&item=53228172&trk=EML_anet_di_pst_ttle
    
    
''    Private Function Excel_QueryTable_Test_3() As Excel.Range
''Method here creates .QueryTable of type xlOLEDBQuery
''http://support.microsoft.com/kb/247412
    On Error GoTo errsub
'
'    'Create the QueryTable
'    Dim ConnString As String
'
'    'Not all extended properties are supported.
'    ConnString = "OLEDB;Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ThisWorkbook.FullName & ";Extended Properties=""Excel 8.0;HDR=YES;IMEX=1;MAXSCANROWS=0"""
'
'    'Delete all previous tables on sheet.
'    While Sheet1.QueryTables.Count > 0
'        Sheet1.QueryTables(1).Delete
'    Wend
'    Sheet1.UsedRange.Delete
'
'    With Sheet1.QueryTables.Add(ConnString, Sheet1.Range("A1"), "SELECT * FROM [Sheet2$]")
'        .RefreshStyle = xlInsertEntireRows
'        .AdjustColumnWidth = False
''        .TextFileColumnDataTypes = Array(xlTextFormat, xlTextFormat, xlTextFormat) 'xlSkipColumn, xlGeneralFormat)
'        .Refresh True 'Displays the table
''        MsgBox .QueryType
'        If .FetchedRowOverflow = True Then
'            'return truncated result set
'            Set Excel_QueryTable_Test_3 = Sheet1.Range("A1").Resize(Sheet1.Rows.Count, Sheet1.Columns.Count)
'        Else
'            Set Excel_QueryTable_Test_3 = .ResultRange
'        End If
'    End With
'    Exit Function
errsub:
    Debug.Print Err.Description
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''UserForm Code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub UserForm_Initialize()   'Private Sub Class_Initialize()
    Dim oPW As New clsPositionWindow
    
    Set m_Application = Application
    Set m_DBConnection = CreateObject("ADODB.Connection")
    Set m_Command = CreateObject("ADODB.Command")
    Me.Hide
    
    'Set default position
    Call oPW.ForceWindowIntoWorkArea(Me, vbStartUpCenterParent)
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
    Dim strTempFile As String
    
    If Not m_DBConnection Is Nothing Then
        If m_DBConnection.State = ObjectStateEnum.adStateOpen Then
            m_DBConnection.Close
        End If
        Set m_DBConnection = Nothing
    End If
    
    Set m_Command = Nothing
    
    Call DeleteTempFile(DataSourceBackupExcelFilePath)
    Call DeleteTempFile(SchemaINIPath)
End Sub

Private Function SaveTempExcelFile(ByRef strFileName As String) As String
    'strFileName only used to determine the Excel file type.
    On Error GoTo errsub
    Dim strPath As String
    Dim strFileEx As String
    
    strFileEx = ParsePath(strFileName, FileExtension)
    
    strPath = GetTempFileName(strFileEx)

'    Application.DisplayAlerts = False
    Call m_Application.ActiveWorkbook.SaveCopyAs(strPath)
'    Application.DisplayAlerts = True

    SaveTempExcelFile = strPath
'    Debug.Print "Saved temporary copy of file connection to: " & strFile
errsub:
End Function

Public Function SaveTempExcelWorksheetAsAsCSV(wksSource As Object) As String 'Object = Excel.Worksheet
    On Error GoTo errsub
    Dim strPath As String
    
    Const xlCSVMSDOS As Long = 24
    
    strPath = GetTempFileName("csv")

'    Application.DisplayAlerts = False
    Call wksSource.SaveAs(strPath, xlCSVMSDOS, False)
'    Application.DisplayAlerts = True

    SaveTempExcelWorksheetAsAsCSV = strPath
'    Debug.Print "Saved temporary copy of file connection to: " & strFile
errsub:
End Function

Private Function SaveWorksheetAsCSV(wksSource As Object, strFileName As String)
'Private Function SaveWorksheetAsCSV(wksSource As Excel.Worksheet, strFileName As String)
    Const xlCSVMSDOS As Long = 24
    Call wksSource.SaveAs(strFileName, xlCSVMSDOS, False)
End Function

Private Function DeleteTempFile(strFilePath As String) As Boolean
'Try to delete file if it exists.
    On Error GoTo errsub
    
    If ParsePath(strFilePath, FileName) <> vbNullString Then
        If DoesFilePathExist(strFilePath) = True Then
            Call VBA.Kill(strFilePath)
        End If
    End If
    DeleteTempFile = True
errsub:
'File in use?
End Function

Public Function ParsePath(ByVal strPath As String, iMode As PathParseMode) As String
    'Take the path passed in and return the filename, or the path base upon iMode.
    'Path returns with trailing Application.PathSeparator ("\")
    'If file name is passed in to returnPath, empty string is returned.

    On Error GoTo errsub

    Dim fso As Object
    Set fso = CreateObject("Scripting.Filesystemobject")
    
    Select Case iMode
        Case PathParseMode.FileExtension
            ParsePath = fso.GetExtensionName(strPath) 'File extension name
        Case PathParseMode.FileName
            ParsePath = fso.GetFileName(strPath) 'File name with extension
        Case PathParseMode.Path
            If VBA.InStr(1, strPath, "\") > 0 Then  '"\" exists
                strPath = fso.GetParentFolderName(strPath) 'File path
                ParsePath = fso.BuildPath(strPath, "\") 'Add "\"
            End If
        Case PathParseMode.FileNameWithoutExtension
            ParsePath = fso.GetBaseName(strPath) 'File name without extension
    End Select
errsub:
    Set fso = Nothing
End Function

Private Function IsFileOpen(strFullFilePath As String) As Boolean
'http://www.xcelfiles.com/IsFileOpen.html
'Check if File is Open.
    Dim hdlFile As Long

    'Error is generated if you try opening a File for ReadWrite lock >> MUST BE OPEN!
    If DoesFilePathExist(strFullFilePath) = True Then
        On Error GoTo errsub:
        hdlFile = FreeFile
'        Open strFullFilePath For Random Access Read Write Lock Read Write As #hdlFile 'Works for Excel files, but not text files.
        Open strFullFilePath For Input Lock Read As #hdlFile 'http://www.cpearson.com/excel/IsFileOpen.aspx
        Close hdlFile
    End If
    
    IsFileOpen = False
    Exit Function
    
errsub: 'Someone has file open
    IsFileOpen = True
'    Close hdlFile
End Function

Private Sub SampleConnectPivotTableWithDataSource()
'Untested sample code to connect Pivot table with datasource.  Can be used to represent data that exceeds Excel limit.
'http://www.sqlservercentral.com/Forums/Topic551557-60-3.aspx
'Sub Recompute(Source As String, P1 As String, P2 As String)
'If P1 = "" Then Exit Sub
'If P2 = "" Then Exit Sub
'
'
'On Error GoTo ErrorHandler
'
'Select Case Source
'Case "Pivot"
'    Range("C15").Select
'
'    With ActiveWorkbook.PivotCaches.Item(1)
'        .Connection = _
'        "ODBC;DRIVER=SQL Server;SERVER=datamart.onsemi.com;UID=myuid;APP=Microsoft Office 2003;WSID=myuid-D4;DATABASE=SM;Trusted_Connection=Yes"
'        .CommandType = xlCmdSql
'         .CommandText = Array("exec asp_World '" & P1 & "', '" & P2 & "'")
'    End With
'Range("A2:L2").FormulaR1C1 = "'ASP@Mix change impact for " & P1 & " To " & P2
'Case "Raw Data"
'  Sheets("raw").Select
'    Range("D10").Select
'    With Selection.QueryTable
'        .Connection = _
'        "ODBC;DRIVER=SQL Server;SERVER=datamart.onsemi.com;UID=myuid;APP=Microsoft Office 2003;WSID=myuid-D4;DATABASE=SM;Trusted_Connection=Yes"
'        .CommandType = xlCmdSql
'        .CommandText = Array("exec asp_World '" & P1 & "', '" & P2 & "'")
'        .Refresh BackgroundQuery:=False
'    End With
'End Select
'
'Exit Sub
'ErrorHandler:
'MsgBox ("An error occured, please try again. " & P1 & ", " & P2)
'End Sub
'
'
'Private Function adoVar() As Integer
'Dim Connstr As String
'Dim cnnConnect As ADODB.Connection
'Dim rstRecordset As ADODB.RecordSet
'Dim MyCmd As String
'' this function requires a reference to ADO active X libraries
'
'MyCmd = "SELECT case when r.Status = 'Ready' then 1 else 0 end AS Status From onglobals.dbo.tb_IsReady r WHERE (r.Name = 'ASP_World_tbl') "
'Connstr = "ODBC;DRIVER=SQL Server;SERVER=rocky;UID=myuid;APP=Microsoft Office 2003;WSID=myuid-D4;DATABASE=ONglobals;Trusted_Connection=Yes" _
'
'Set cnnConnect = New ADODB.Connection
'cnnConnect.Open Connstr
'
'Set rstRecordset = New ADODB.RecordSet
'rstRecordset.Open Source:=MyCmd, ActiveConnection:=cnnConnect, _
'    CursorType:=adOpenDynamic, LockType:=adLockReadOnly, Options:=adCmdText
'adoVar = rstRecordset.Fields(0).value
'
'cnnConnect.Close
'Set rstRecordset = Nothing
'
'End Function
End Sub

Public Function GetCustomDocumentProperty(docActive As Object, strPropertyName As String, Optional bIsCustomProperty As Boolean = True) As Variant
'http://www.cpearson.com/excel/docprop.htm 'Modified/Simplified
'Check for empty return with IsEmpty().
    On Error Resume Next
    
    If bIsCustomProperty = True Then
        GetCustomDocumentProperty = docActive.CustomDocumentProperties(strPropertyName).Value
    Else
        GetCustomDocumentProperty = docActive.BuiltinDocumentProperties(strPropertyName).Value
    End If

End Function

Public Sub SetCustomDocumentProperty(docActive As Object, strPropertyName As String, vPropertyValue As Variant, Optional bIsCustomProperty As Boolean = True)
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
        Set DocProps = docActive.CustomDocumentProperties
    Else
        Set DocProps = docActive.BuiltinDocumentProperties
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
    If GetCustomDocumentProperty(docActive, strPropertyName, bIsCustomProperty) = Empty Then
        Call DocProps.Add(strPropertyName, False, TheType, vPropertyValue)
    End If
    
    DocProps(strPropertyName).Value = vPropertyValue
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'WMI Code - Windows Management Instrumentation
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetLocalMachineNameWMI() As String
'http://www.microsoft.com/technet/scriptcenter/resources/qanda/apr06/hey0425.mspx
    Dim objWMIService As Object 'SWbemServicesEx
    Dim objServiceSet As Object 'SWbemObjectEx
    '    Dim objService As Object 'SWbemObjectSet
    Dim errReturnCode As Long
    Dim obj As Object 'SWbemObjectEx
    
    On Error GoTo errsub
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    Set objServiceSet = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
    For Each obj In objServiceSet
            GetLocalMachineNameWMI = obj.name
            Exit Function
    Next obj
    Exit Function
errsub:
End Function

Private Function ReStartServiceForTimeWMI(ByVal strServerName As String, ByVal strServiceName, ByVal lTimeoutSeconds As Long) As Boolean
    On Error GoTo errsub

    'Attempt to start service if not running
    If GetServiceStateByNameWMI(strServerName, strServiceName) <> "Stopped" Then
        Call StopServiceForTimeWMI(strServerName, strServiceName, lTimeoutSeconds)
    End If
    
    If GetServiceStateByNameWMI(strServerName, strServiceName) = "Stopped" Then
        If StartServiceForTimeWMI(strServerName, strServiceName, lTimeoutSeconds) Then
            ReStartServiceForTimeWMI = True
        End If
    End If

    Exit Function
errsub:
    ReStartServiceForTimeWMI = False
End Function

Private Function StartServiceForTimeWMI(ByVal strServerName As String, ByVal strServiceName, ByVal lTimeoutSeconds As Long) As Boolean
    On Error GoTo errsub
    Dim dateCurrentTime As Date
    Dim lRet As Long
    
    dateCurrentTime = Now

    'Attempt to start service if not running
    If GetServiceStateByNameWMI(strServerName, strServiceName) <> "Running" Then
        If StartServiceByNameWMI(strServerName, strServiceName) = 0 Then    'Start was successful.
            Do Until GetServiceStateByNameWMI(strServerName, strServiceName) = "Running"
                DoEvents
                If VBA.DateDiff("s", dateCurrentTime, Now) > lTimeoutSeconds Then 'Or GetServiceStateByNameWMI(strServerName, strServiceName) = "Running" Then
                    StartServiceForTimeWMI = False
                    Exit Function
                End If
            Loop
        End If
    End If
    StartServiceForTimeWMI = True
    Exit Function
errsub:
    StartServiceForTimeWMI = False
End Function

Private Function StopServiceForTimeWMI(ByVal strServerName As String, ByVal strServiceName, ByVal lTimeoutSeconds As Long) As Boolean
    On Error GoTo errsub
    Dim dateCurrentTime As Date
    Dim lRet As Long
    
    dateCurrentTime = Now

    'Attempt to stop service if running
    If GetServiceStateByNameWMI(strServerName, strServiceName) <> "Stopped" Then
        If StopServiceByNameWMI(strServerName, strServiceName) = 0 Then    'Stop was successful.
            Do Until GetServiceStateByNameWMI(strServerName, strServiceName) = "Stopped"
                DoEvents
                If VBA.DateDiff("s", dateCurrentTime, Now) > lTimeoutSeconds Then 'Or GetServiceStateByNameWMI(strServerName, strServiceName) = "Running" Then
                    StopServiceForTimeWMI = False
                    Exit Function
                End If
            Loop
        End If
    End If
    StopServiceForTimeWMI = True
    Exit Function
errsub:
    StopServiceForTimeWMI = False
End Function

Private Function StartServiceByNameWMI(ByVal strServerName As String, ByVal strServiceName As String) As Long
'http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/services/#StopService.htm
    Dim objWMIService As Object 'SWbemServicesEx
    Dim objServiceSet As Object 'SWbemObjectSet
'    Dim objService As SWbemObjectSet
    Dim errReturnCode As Long
    Dim obj As Object 'SWbemObjectEx
    
    errReturnCode = -1
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServerName & "\root\cimv2")
    Set objServiceSet = objWMIService.ExecQuery("Select * from win32_Service Where Name = '" & strServiceName & "'") 'One service

    For Each obj In objServiceSet
'        If obj.Name = strServiceName Then
'            Debug.Print obj.Name & " state is: " & obj.State
            If obj.State = "Running" Then
                errReturnCode = 0
            ElseIf obj.State = "Paused" Then
                errReturnCode = obj.ResumeService
            Else
                errReturnCode = obj.StartService    '0 = Success, 10 if already started.    obj.State = "Running"
    '            errReturnCode = obj.StopService     '0 = Success, 5 if already started stopped. obj.State = "Stopped"
    '            errReturnCode = obj.ResumeService   '0 = Success(was paused), 6 = Currently stopped, 10 = was already running.
        'PauseService & ResumeService
            End If
            Exit For
'        End If
    Next obj
    
    StartServiceByNameWMI = errReturnCode
End Function

Private Function StopServiceByNameWMI(ByVal strServerName As String, ByVal strServiceName As String) As Long
'http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/services/#StopService.htm
    Dim objWMIService As Object 'SWbemServicesEx
    Dim objServiceSet As Object 'SWbemObjectSet
'    Dim objService As SWbemObjectSet
    Dim errReturnCode As Long
    Dim obj As Object 'SWbemObjectEx
    
    errReturnCode = -1
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServerName & "\root\cimv2")
    Set objServiceSet = objWMIService.ExecQuery("Select * from win32_Service Where Name = '" & strServiceName & "'") 'One service

    For Each obj In objServiceSet
'        If obj.Name = strServiceName Then
'            Debug.Print obj.Name & " state is: " & obj.State
            If obj.State = "Stopped" Then
                errReturnCode = 0
'            ElseIf obj.State = "Paused" Then
'                errReturnCode = obj.ResumeService
            Else
'                errReturnCode = obj.StartService    '0 = Success, 10 if already started.    obj.State = "Running"
                errReturnCode = obj.StopService     '0 = Success, 5 if already started stopped. obj.State = "Stopped"
    '            errReturnCode = obj.ResumeService   '0 = Success(was paused), 6 = Currently stopped, 10 = was already running.
        'PauseService & ResumeService
            End If
            Exit For
'        End If
    Next obj
    
    StopServiceByNameWMI = errReturnCode
End Function

Private Function GetServiceStateByNameWMI(ByVal strServerName As String, ByVal strServiceName As String) As String
'http://www.activexperts.com/activmonitor/windowsmanagement/adminscripts/services/#StopService.htm
    Dim objWMIService As Object 'SWbemServicesEx
    Dim objServiceSet As Object 'SWbemObjectEx
    '    Dim objService As Object 'SWbemObjectSet
    Dim errReturnCode As Long
    Dim obj As Object 'SWbemObjectEx
    
    On Error GoTo errsub
    
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strServerName & "\root\cimv2")
    Set objServiceSet = objWMIService.ExecQuery("Select * from win32_Service Where Name = '" & strServiceName & "'") 'One service
    For Each obj In objServiceSet
'        If obj.Name = strServiceName Then
            GetServiceStateByNameWMI = obj.State
'            Debug.Print obj.Name & " state is: " & obj.State
'            Debug.Print obj.Name & " Status is: " & obj.Status
            Exit Function
'        End If
    Next obj
    Exit Function
errsub:
    GetServiceStateByNameWMI = "Error"
End Function

'=======================================================================================================
'Helper Functions
'=======================================================================================================
Private Function DoesFilePathExist(strPath As String, Optional IsDirectory As Boolean = False) As Boolean
'Replacement for Dir() function that can return error #52 "Bad file name or number" when testing for invalid file with a network path.
    On Error GoTo errsub:
    
    Dim fso As Object 'FileSystemObject
     
    Set fso = CreateObject("Scripting.Filesystemobject")
    
    If IsDirectory = False Then
        DoesFilePathExist = fso.FileExists(strPath)
    Else
        DoesFilePathExist = fso.FolderExists(strPath)
    End If
errsub:
    Set fso = Nothing
End Function

Private Function GetServerInstanceParts(ByVal strSource As String, ByRef strServerName As String, ByRef strInstanceName As String) As Boolean
'Pass backserver name and instance name based on strSource
    strSource = Replace(strSource, "\\", Empty) 'Sometimes on front of string
    If VBA.InStr(1, strSource, "\") <> 0 Then 'Server Instance Used
        strServerName = UCase(Left(strSource, VBA.InStr(1, strSource, "\") - 1))
        strInstanceName = UCase(Mid(strSource, VBA.InStr(1, strSource, "\") + 1))
        GetServerInstanceParts = True
    Else
        strServerName = strSource
    End If

'Can use these to get same information from active DB connection.
'http://msdn.microsoft.com/en-us/library/ms174396.aspx
'    Execute "SELECT SERVERPROPERTY('InstanceName')"
'    Execute "SELECT SERVERPROPERTY('MachineName')"
'    Execute "SELECT SERVERPROPERTY('ServerName')"
End Function

Private Function GetValidSQLDBName(DBDataProvider As eDataProvider, ByVal strDataSource As String, ByRef strDataTable As String, TrustedConnection As Boolean, Optional bAllowSelection As Boolean = True, Optional bOnlyShowNormalizedSimbryoDBNames As Boolean = False) As Boolean
''Validates name passed in against list in default DB 'master' DB.
'Replace with call to m_DBConnection.OpenSchema()?
''If not found, puts list of databases currently available in userform for selection.
''Sets strDataTableto new value if found.
'    On Error GoTo errsub
'
    Dim strDBList() As String
    Dim strExec As String
    Dim lCursor As Long
    Dim strCommandText As String
    Dim oRecordSet As Object    'ADODB.Recordset
    Dim objDB As clsDBConnectivity
    
    Set objDB = New clsDBConnectivity

    If ConnectToDB(DBDataProvider, strDataSource, "master", strDataTable, Empty, Empty, TrustedConnection) = True Then
        strExec = "SELECT name FROM master..sysdatabases"
'        strExec = "EXECUTE sp_databases" 'TSQL stored proc
        Set oRecordSet = ExecuteRecordset(strExec)

        'Get list of DBs from master DB
        Do While Not oRecordSet.EOF
            If oRecordSet("name") <> "master" And oRecordSet("name") <> "model" And _
                    oRecordSet("name") <> "msdb" And oRecordSet("name") <> "tempdb" Then
'            If oRecordset("DATABASE_NAME") <> "master" And oRecordset("DATABASE_NAME") <> "model" And _
'                    oRecordset("DATABASE_NAME") <> "msdb" And oRecordset("DATABASE_NAME") <> "tempdb" Then
                If IsArrayEmpty(strDBList) Then
                    ReDim strDBList(0) As String
                Else
                    ReDim Preserve strDBList(UBound(strDBList) + 1) As String
                End If
                strDBList(UBound(strDBList)) = oRecordSet("name")
'                strDBList(UBound(strDBList)) = oRecordset("DATABASE_NAME")
'                Debug.Print oRecordset("DATABASE_NAME")
                If oRecordSet("name") = strDataTable Then
'                If oRecordset("DATABASE_NAME") = strDataTable Then
                    GetValidSQLDBName = True
                    Exit Do
                End If
            End If
            oRecordSet.MoveNext
        Loop

        If IsArrayEmpty(strDBList) = False And GetValidSQLDBName = False And bAllowSelection = True Then
            Me.Caption = "Database Connection"
            Me.frmDBSelection.Caption = "Available Databases:"
            Me.cmbDataTableSelection.ControlTipText = "Supplied database is invalid. Select database to open from the available list."
            
            Call AddDBComboBoxNames(Me.cmbDataTableSelection, strDBList, bOnlyShowNormalizedSimbryoDBNames)
'            If bOnlyShowSimbryoDBNames Then
'                Me.cmbDBSelection.List() = FixDBNames(strDBList) 'Need to pass back original name as a result.
'            Else
'                Me.cmbDBSelection.List() = strDBList
'            End If
            Me.cmbDataTableSelection.ListIndex = 0 'set to first item in list.

            If m_Application.name = "Microsoft Excel" Then
                lCursor = m_Application.Cursor
                m_Application.Cursor = m_Application.XlMousePointer.xlDefault
            End If

            Me.Show vbModal 'Unload of form canceled with QueryClose so that we don't lose object references.

            'Pass real DB name back by reference (not always what was visible in dropdown).
            strDataTable = Me.cmbDataTableSelection.List(Me.cmbDataTableSelection.ListIndex, 1)
            GetValidSQLDBName = True
        End If
    End If

errsub: 'Fall through intentional
    If m_Application.name = "Microsoft Excel" Then
        m_Application.Cursor = lCursor
    End If
    
    If Not oRecordSet Is Nothing Then
        If oRecordSet.State = ObjectStateEnum.adStateOpen Then
            oRecordSet.Close
        End If
        Set oRecordSet = Nothing
    End If
    
    Set objDB = Nothing
    
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Private Function GetUsedRowByStartCell(ByVal rngInput As Object) As Object
'Private Function GetUsedRowByStartCell(ByVal rngInput As Excel.Range) As Excel.Range
'Returns range of used row, starting with first cell in rngInput.
'Returns nothing if no range.
    On Error GoTo errsub
    Dim rng As Object 'Excel.Range

    Set rngInput = rngInput(1)
    Set rng = GetLastUsedCellInRowByStartCell(rngInput) 'Checks if found cell is before rngInput
    If Not rng Is Nothing Then
        Set GetUsedRowByStartCell = rngInput.Parent.Range(rngInput, rng)
    End If
    
errsub:
If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Private Function GetLastUsedCellInRowByStartCell(rngInput As Object) As Object
'Private Function GetLastUsedCellInRowByStartCell(rngInput As Excel.Range) As Excel.Range
    'If row is empty or empty at and after start cell, passes back nothing
    Dim rng As Object 'Excel.Range
    
    Set rngInput = rngInput(1)
    Set rng = GetLastUsedCellInRowByValue(rngInput)
    If Not rng Is Nothing Then
        Set GetLastUsedCellInRowByStartCell = rng
    End If
End Function

Private Function GetLastUsedCellInRowByValue(rngStart As Object, Optional vValue As Variant = "*") As Object
'Private Function GetLastUsedCellInRowByValue(rngStart As Excel.Range, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by rows.
    Dim rng As Object 'Excel.Range
    
    Const xlValues = -4163 'Excel.XlFindLookIn.xlValues
    Const xlWhole = 1 'Excel.XlLookAt.xlWhole
    Const xlByColumns = 2 'Excel.XlSearchOrder.xlByColumns
    Const xlPrevious = 2 'Excel.XlSearchDirection.xlPrevious
    
    Set rng = rngStart.EntireRow.Find(vValue, , xlValues, xlWhole, xlByColumns, xlPrevious)
    If rng.Row = rngStart.Row Then 'Will return different column if searching a blank area.
        Set GetLastUsedCellInRowByValue = rng
    End If
End Function

Function GetLastUsedCell(wksSource As Object) As Object
'Function GetLastUsedCell(wksSource As Excel.Worksheet) As Excel.Range
    'Error-handling here in case there is not any data in the worksheet
    On Error GoTo errsub:
    Set GetLastUsedCell = wksSource.Cells(GetLastUsedRow(wksSource).Row, GetLastUsedColumn(wksSource).Column)
errsub:
End Function

Private Function GetLastUsedRow(wksSource As Object, Optional vValue As Variant = "*") As Object
'Private Function GetLastUsedRow(wksSource As Excel.Worksheet, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by Rows.
    On Error GoTo errsub:
    
    Const xlValues = -4163 'Excel.XlFindLookIn.xlValues
    Const xlWhole = 1 'Excel.XlLookAt.xlWhole
    Const xlByRows = 1 'Excel.XlSearchOrder.xlByRows
    Const xlPrevious = 2 'Excel.XlSearchDirection.xlPrevious
    
    Set GetLastUsedRow = wksSource.UsedRange.Cells.Find(vValue, , xlValues, xlWhole, xlByRows, xlPrevious)
errsub:
End Function

Private Function GetLastUsedColumn(wksSource As Object, Optional vValue As Variant = "*") As Object
'Private Function GetLastUsedColumn(wksSource As Excel.Worksheet, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by Columns.
    On Error GoTo errsub:
    
    Const xlValues = -4163 'Excel.XlFindLookIn.xlValues
    Const xlWhole = 1 'Excel.XlLookAt.xlWhole
    Const xlByColumns = 2 'Excel.XlSearchOrder.xlByColumns
    Const xlPrevious = 2 'Excel.XlSearchDirection.xlPrevious
    
    Set GetLastUsedColumn = wksSource.UsedRange.Cells.Find(vValue, , xlValues, xlWhole, xlByColumns, xlPrevious)
errsub:
End Function

Private Function AddDBComboBoxNames(ByRef combo As ComboBox, strDBList() As String, Optional OnlyShowSimbryoDBNames As Boolean = False) As String()
'Pass in string array of names to go into combo.  Combo contains display name in (0) and holds real name in (1).
'If OnlyShowSimbryoDBNames = True then Simbryo format is changed to readable format.
'Raw Simbryo DB format is: <Session Name>#qqq#<Product Name>#qqq#<Version>
    Dim lCount As Long
    Dim arryTemp As Variant
    
    combo.Clear
    For lCount = 0 To UBound(strDBList)
        On Error Resume Next    'Can error if unexpected format is found.
        'Raw Simbryo DB format is: <Session Name>#qqq#<Product Name>#qqq#<Version>
        If OnlyShowSimbryoDBNames Then
            arryTemp = VBA.Split(strDBList(lCount), "#qqq#") 'Simbryo names have #qqq# to seperate fields.
            If UBound(arryTemp) = 2 Then  'Simbryo DB name
                'Raw Simbryo DB format is:
                '<Session Name>#qqq#<Product Name>#qqq#<Version>#qqq#
                Call combo.AddItem(arryTemp(1) & " - " & arryTemp(0) & " - " & Replace(arryTemp(2), "_", "."))       'Normal Name
                combo.List(combo.ListCount - 1, 1) = strDBList(lCount)   'Add Real Name
            End If
        Else 'Uncorrected formats
            Call combo.AddItem(strDBList(lCount))       'Normal Name
            combo.List(combo.ListCount - 1, 1) = strDBList(lCount)   'Add Real Name
        End If
    Next lCount
End Function

Public Function IsArrayEmpty(checkArray As Variant) As Boolean
'Used to tell is array is empty.
'Uninitialized arrays or arrays that are cleared with Erase().
    Dim lngTmp As Long
    On Error GoTo emptyError

    'Here is where it happens.
    'If you haven't used Redim on your array
    'Ubound will return an error
    lngTmp = UBound(checkArray)
    If lngTmp = -1 Then GoTo emptyError 'Returns -1 if empty.
    IsArrayEmpty = False
    Exit Function
emptyError:
    Err.Clear 'Clear out error code
    On Error GoTo 0 'Turn off error checking
    IsArrayEmpty = True
End Function

'Public Function GetNumberOfArrayDimensions(Arr As Variant) As Long
'    'http://www.cpearson.com/Excel/VBAArrays.htm
'    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    ' NumberOfArrayDimensions
'    ' This function returns the number of dimensions of an array. An unallocated dynamic array
'    ' has 0 dimensions. This condition can also be tested with IsArrayEmpty.
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    On Error Resume Next
'    Dim Ndx As Integer
'    Dim Res As Integer
'
'    ' Loop, increasing the dimension index Ndx, until an error occurs.
'    ' An error will occur when Ndx exceeds the number of dimension
'    ' in the array. Return Ndx - 1.
'    Do
'        Ndx = Ndx + 1
'        Res = UBound(Arr, Ndx)
'    Loop Until Err.Number <> 0
'
'    GetNumberOfArrayDimensions = Ndx - 1
'    Err.Clear
'End Function

Private Function CollectionToArray(col As Collection) As Variant
    Dim lCount As Long
    Dim vArray As Variant
    Dim vItem As Variant
    
    ReDim vArray(col.Count - 1) As Variant
    
    For Each vItem In col
        vArray(lCount) = vItem
        lCount = lCount + 1
    Next
    
    CollectionToArray = vArray
End Function

Public Function TransposeArray(vArray As Variant) As Variant
'Custom Function to Transpose vArray
'Use instead of Excel's WorksheetFunction.Transpose() due to limitations with call listed at link.
'http://support.microsoft.com/kb/246335
    Dim X As Long
    Dim Y As Long
    Dim Xupper As Long
    Dim Yupper As Long
    Dim XLower As Long
    Dim YLower As Long
    Dim TempArray As Variant

    Yupper = UBound(vArray, 1)
    YLower = LBound(vArray, 1)
    Select Case NumberOfArrayDimensions(vArray)
        Case 1
            ReDim TempArray(0 To 0, YLower To Yupper)
            For Y = YLower To Yupper
                TempArray(0, Y) = vArray(Y)
            Next
        Case 2
            Xupper = UBound(vArray, 2)
            XLower = LBound(vArray, 2)
            ReDim TempArray(XLower To Xupper, YLower To Yupper)
            For X = XLower To Xupper
                For Y = YLower To Yupper
                    TempArray(X, Y) = vArray(Y, X)
                Next
            Next
    End Select

    TransposeArray = TempArray
End Function

Private Function NumberOfArrayDimensions(vArray As Variant) As Long
'Returns the number of dimensions in an array
'http://support.microsoft.com/kb/152288
    On Error GoTo FinalDimension
    
    Dim DimNum As Long
    Dim ErrorCheck As Long

    'Visual Basic for Applications arrays can have up to 60000 dimensions.
    For DimNum = 1 To 60000
        ErrorCheck = LBound(vArray, DimNum)
    Next

    Exit Function
    
FinalDimension:
    NumberOfArrayDimensions = DimNum - 1
End Function

Private Function GetOSVersionNumber() As Single
'http://msdn.microsoft.com/en-us/library/ms724834(v=vs.85).aspx
'6.1 = Windows 7
'6.0 = Windows Vista
'5.1 - Windows XP
    Dim OSV As OSVERSIONINFO

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) <> False Then
        GetOSVersionNumber = VBA.CSng(OSV.dwVerMajor & "." & OSV.dwVerMinor)
    End If
End Function

Private Function GetTempFileName(Optional strFileExtension As String = "tmp") As String
'When called using UNIQUE_NAME creates unique temp file name.
'WinAPI GetTempPath() can also be used to get temp path instead of VBA.Environ("TEMP").
'Returns full path name to unique file.
    On Error GoTo errsub
    Dim strResult As String
    
    strResult = VBA.Space(MAX_PATH)
    Call GetTempFileNameA(VBA.Environ("TEMP"), "TMP", UNIQUE_NAME, strResult)
    strResult = VBA.Left(strResult, VBA.InStr(strResult, VBA.Chr(0)) - 1)
    If DoesFilePathExist(strResult) = True Then 'File creation ensures unique file name.
        Call VBA.Kill(strResult)
        strFileExtension = VBA.Replace(strFileExtension, ".", "")
        GetTempFileName = VBA.Left(strResult, VBA.Len(strResult) - 3) & strFileExtension
    End If
errsub:
End Function

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

Public Function GetWorkbookProperty(wkbActive As Excel.Workbook, strPropertyName As String, Optional bIsCustomProperty As Boolean = True) As Variant
'http://www.cpearson.com/excel/docprop.htm 'Modified/Simplified
'Check for empty return with IsEmpty().
    On Error Resume Next

    If bIsCustomProperty = True Then
        GetWorkbookProperty = wkbActive.CustomDocumentProperties(strPropertyName).Value
    Else
        GetWorkbookProperty = wkbActive.BuiltinDocumentProperties(strPropertyName).Value
    End If

End Function

Public Sub SetWorkbookProperty(wkbActive As Excel.Workbook, strPropertyName As String, vPropertyValue As Variant, Optional bIsCustomProperty As Boolean = True)
'Private Sub SetProperty(WorkbookName As String, PropName As String, PValue As Variant, PropCustom As Boolean)
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
