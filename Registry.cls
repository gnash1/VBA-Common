VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API_Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'//////////////////////////////////////////////////////////////////////
'//
'// CLASS:      API_Registry.cls
'// AUTHOR:     Josh Holdaway
'// PURPOSE:    To read/write values inside the windows registry
'// PROBLEMS:   None
'//
'//  HISTORY   ////////////////////////////////////////////////////////
'//
'// Created  January 28, 2002 : Josh Holdaway
'//
'//////////////////////////////////////////////////////////////////////

'declare our user-defined type's
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Boolean
End Type

'declare our enumerations
Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Enum lpType
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_BINARY = 3
    REG_DWORD = 4
End Enum

'declare our registry functions
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function ExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long

'define our registry events
Public Event onSuccess()
Public Event onFailed()

Private Const KEY_ALL_ACCESS = &HF003F
Private Const READ_CONTROL As Long = &H20000
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_SET_VALUE As Long = &H2
Private Const KEY_CREATE_SUB_KEY As Long = &H4
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const STANDARD_RIGHTS_WRITE As Long = (READ_CONTROL)
Private Const KEY_READ As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Private Const KEY_WRITE As Long = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

'declare our private variables
Private HKeySet As Long
Private DType As lpType
Private KPath As String
Private Handle As Long
Private success As Long
Private ValueN As String
Private sData As String
Private Bdata() As Byte
Private DData As Long
Private Buffsize As Long
Private sec As SECURITY_ATTRIBUTES

Public Property Get hKey() As hKey
    hKey = HKeySet
End Property

Public Property Let hKey(ByVal Key As hKey)
    HKeySet = Key
End Property

Public Property Get DataType() As lpType
    DataType = DType
End Property

Public Property Let DataType(ByVal Dat As lpType)
    DType = Dat
End Property

Public Property Get SubKey() As String
    SubKey = KPath
End Property

Public Property Let SubKey(ByVal KP As String)
    KPath = KP
End Property

Public Property Get isSuccess() As Boolean
    
    'is there a success?
    If success = 0 Then
        isSuccess = True
    Else
        isSuccess = False
    End If
    
    success = 1
End Property

Public Property Get ValueName() As String
    ValueName = ValueN
End Property

Public Property Let ValueName(ByVal val As String)
    ValueN = val
End Property

Public Property Get Data() As Variant
    'return the data from the registry
    Select Case DType
        Case REG_SZ
            Data = sData
        Case REG_EXPAND_SZ
            Data = sData
        Case REG_BINARY
            Data = Bdata
        Case REG_DWORD
            Data = DData
        Case Else
            Data = sData
    End Select
End Property

Public Property Let Data(ByVal Dat1 As Variant)
    'store the data, so we can write it to the registry
    Select Case DType
        Case REG_SZ
            sData = Dat1 & vbNullChar
        Case REG_EXPAND_SZ
            sData = Dat1 & vbNullChar
        Case REG_BINARY
            Buffsize = Len(Dat1) / 2
            ReDim Bdata(Buffsize) As Byte
            Bdata = Dat1
        Case REG_DWORD
            Buffsize = 4
            DData = Dat1
    End Select
End Property

Public Sub DeleteValue()

    success = 0 'set initial value to zero
    
    'try to open the key first
    success = RegOpenKeyEx(HKeySet, KPath, 0, KEY_ALL_ACCESS, Handle)
    If success <> 0 Then GoTo Failed 'if failed, exit
    
    'try to delete the value
    success = RegDeleteValue(Handle, ValueN)
    If success <> 0 Then GoTo Failed 'if failed, exit
    
    'now close the key
    success = RegCloseKey(Handle)
    If success <> 0 Then GoTo Failed
    
    RaiseEvent onSuccess
    Exit Sub
    
Failed:
    success = RegCloseKey(Handle)
    RaiseEvent onFailed
End Sub

Public Sub DeleteKey()

    'try to delete the key
    success = RegDeleteKey(HKeySet, KPath)
    If success <> 0 Then GoTo Failed 'did it fail?
    
    RaiseEvent onSuccess
    Exit Sub
    
Failed:
    success = 1
    RaiseEvent onFailed
End Sub

Public Sub CreateKey()

    'declare our variables
    Dim Neworused As Long
    
    'set our security attributes
    sec.nLength = Len(secattr)
    sec.lpSecurityDescriptor = 0
    sec.bInheritHandle = True
    
    'try to create the key
    success = RegCreateKeyEx(HKeySet, KPath, 0, "", 0, KEY_ALL_ACCESS, sec, Handle, Neworused)
    If success <> 0 Then GoTo Failed 'did it succeed?
    
    'close the key handle
    success = RegCloseKey(Handle)
    If success <> 0 Then GoTo Failed
    RaiseEvent onSuccess
    
    Exit Sub
Failed:
    success = RegCloseKey(Handle)
    success = 1
    RaiseEvent onFailed
End Sub

Public Sub SaveSetting()

    'declare our variable
    Dim Neworused As Long
    
    'set our security attributes
    sec.nLength = Len(secattr)
    sec.lpSecurityDescriptor = 0
    sec.bInheritHandle = True
    
    'try to create our key handle
    success = RegCreateKeyEx(HKeySet, KPath, 0, "", 0, KEY_ALL_ACCESS, sec, Handle, Neworused)
    If success <> 0 Then GoTo Failed 'did it succeed?
    
    'which data type is this?
    Select Case DType
        Case REG_SZ
            'try to set the value
            success = RegSetValueEx(Handle, ValueN, 0, DType, ByVal sData, Len(sData) + 1)
            If success <> 0 Then GoTo Failed
        Case REG_EXPAND_SZ
            'try to set the value
            success = RegSetValueEx(Handle, ValueN, 0, DType, ByVal sData, Len(sData) + 1)
            If success <> 0 Then GoTo Failed
        Case REG_BINARY
            success = RegSetValueEx(Handle, ValueN, 0, DType, Bdata(0), UBound(Bdata()))
            If success <> 0 Then GoTo Failed
        Case REG_DWORD
            success = RegSetValueEx(Handle, ValueN, 0, DType, DData, 4)
            If success <> 0 Then GoTo Failed
    End Select
    
    'close the registry handle
    success = RegCloseKey(Handle)
    If success <> 0 Then GoTo Failed

    RaiseEvent onSuccess
    Exit Sub
Failed:
    success = RegCloseKey(Handle)
    success = 1
    RaiseEvent onFailed
End Sub

Public Sub GetSetting()
    'define our variable buffer
    Dim buff As Long
    
    'open a handle to the key
    success = RegOpenKeyEx(HKeySet, KPath, 0, KEY_READ, Handle)
    If success <> 0 Then GoTo Failed 'succeed?
    
    'which data type?
    Select Case DType
    Case REG_SZ  'String
        success = RegQueryValueEx(Handle, ValueN, 0, DType, "", buff)
        sData = Space(buff)
        success = RegQueryValueEx(Handle, ValueN, 0, DType, ByVal sData, buff)
        'sData = Left$(sData, buff - 1) 'remove chr(0)
        sData = Replace(sData, Chr(0), "")
        
        If success <> 0 Then GoTo Failed
    
    Case REG_EXPAND_SZ
        Dim strTemp As String
        success = RegQueryValueEx(Handle, ValueN, 0, DType, "", buff)
        strTemp = Space(buff)
        success = RegQueryValueEx(Handle, ValueN, 0, DType, ByVal strTemp, buff)
        If success <> 0 Then GoTo Failed
        
        'sData = Space(2 ^ 16)
        success = ExpandEnvironmentStrings(strTemp, "", 0)
        sData = Space(success)
        success = ExpandEnvironmentStrings(strTemp, sData, Len(sData))
        
        sData = Trim(Replace(sData, Chr(0), ""))
        
    Case REG_BINARY 'Binary
        ReDim Bdata(255)
        success = RegQueryValueEx(Handle, ValueN, 0, DType, Bdata(0), buff)
        ReDim Bdata(buff)
        success = RegQueryValueEx(Handle, ValueN, 0, DType, Bdata(0), buff)
        If success <> 0 Then GoTo Failed
        
    Case REG_DWORD 'Dword
        buff = 4
        success = RegQueryValueEx(Handle, ValueN, 0, DType, DData, buff)
        If success <> 0 Then GoTo Failed
    End Select
    
    'close the handle to the key
    success = RegCloseKey(Handle)
    If success <> 0 Then GoTo Failed
    
    RaiseEvent onSuccess
    Exit Sub
Failed:
    success = RegCloseKey(Handle)
    success = 1
    RaiseEvent onFailed
End Sub

Public Function CanWrite() As Boolean

    CanWrite = False

    'open a handle to the key
    success = RegOpenKeyEx(HKeySet, KPath, 0, KEY_WRITE, Handle)
    If success = 0 Then CanWrite = True
    
End Function

