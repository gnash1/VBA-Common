VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File:   clsXML
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        1/2/13
'
'To Do:
'
'Summary:
'Functionality for working with XML files.
'
'Online References:
'http://msdn.microsoft.com/en-us/library/aa468547.aspx -A Beginner's Guide to the XML DOM
'
'
'Notes:
'xmlDoc.getElementsByTagName("Data") 'To get item by name.
'Xquery used to Query XML documents.

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Enum ReadyState
    Uninitialized = 0
    Loading = 1
    Loaded = 2
    Interactive = 3
    Completed = 4
End Enum

'Registry Key / Hive
Private Enum hKey
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002 'WMI Default
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

Private Enum lpType
    REG_SZ = 1
    REG_EXPAND_SZ = 2
'    REG_BINARY = 3
    REG_DWORD = 4
End Enum

Public Function WriteXMLtoRange(XMLFile As String, rngTarget As Variant, Optional ClearDestinationWks As Boolean = False) As Variant 'Excel.Range
    Dim aryResult As Variant
    Dim colItems As Collection
    
    Set rngTarget = rngTarget(1)
    
    If ClearDestinationWks = True Then
        Call rngTarget.Parent.UsedRange.Delete(xlUp)
    End If
    
    aryResult = ExcelXMLSpreadsheetToArray(XMLFile)
    
    rngTarget.Resize(UBound(aryResult, 1) + 1, UBound(aryResult, 2) + 1).Value = aryResult
    Set WriteXMLtoRange = rngTarget
End Function
 
Public Function ExcelXMLSpreadsheetToArray(XMLFile As String) As Variant 'Excel.Range
    On Error GoTo errsub
    'Microsoft XML v6.0 - 'C:\Windows\System32\msxml6.dll
    'Documentation recommends using v3 or v6 only.
        
    'JScript Syntax
    'setProperty - http://msdn.microsoft.com/en-us/library/windows/desktop/ms760290(v=vs.85).aspx
    'Second Level Properties - http://msdn.microsoft.com/en-us/library/windows/desktop/ms766391(v=vs.85).aspx
'    Dim xmlDoc As MSXML2.DOMDocument60
    Dim xmldoc As MSXML2.DOMDocument60 'To parse XML
    Dim xeRows As Object    'MSXML2.IXMLDOMSelection
    Dim xeRow As Object     'MSXML2.IXMLDOMElement 'XML navigation
    Dim xeCell As Object    'MSXML2.IXMLDOMElement 'XML navigation
    Dim aryOutput As Variant
    Dim iRow As Long
    Dim iCol As Long
    Dim sXPath As String
    Dim oXMLNodeList As IXMLDOMNodeList
    Dim bReady As Boolean
    
    sXPath = "//Worksheet/Table[@SecPrice>100]" 'XPath Query
    
    If IsEmpty(GetRegistryValueWMI(REG_SZ, HKEY_CLASSES_ROOT, "CLSID\{2933BF90-7B36-11D2-B20E-00C04F983E60}\VersionList", "6.0")) = False Then
        Set xmldoc = CreateObject("Msxml2.DOMDocument.6.0") 'MSXML2.DOMDocument60 'To parse XML
        Call xmldoc.setProperty("SelectionLanguage", "XPath") 'v6.0
    Else 'Deafult to 3.0
        Set xmldoc = CreateObject("Msxml2.DOMDocument.3.0") 'MSXML2.DOMDocument30 'To parse XML
        Call xmldoc.setProperty("SelectionLanguage", "XSLPattern") 'v3.0
    End If
    
'    xmlDoc.async = False 'http://msdn.microsoft.com/en-us/library/windows/desktop/ms761398(v=vs.85).aspx
    xmldoc.validateOnParse = False 'http://msdn.microsoft.com/en-us/library/windows/desktop/ms762791(v=vs.85).aspx
    
    If xmldoc.Load(XMLFile) Then
        If IsXMLDocReady(xmldoc) = True Then
            Set oXMLNodeList = xmldoc.SelectNodes(sXPath)
            
            Call xmldoc.setProperty("SelectionNamespaces", "xmlns:ns='urn:schemas-microsoft-com:office:spreadsheet'") 'Specify defualt namespace with alias set to 'ns'
            Set xeRows = xmldoc.SelectNodes("//ns:Row") 'Selection
            ReDim aryOutput(xeRows.Length - 1, xmldoc.SelectNodes("//ns:Column").Length - 1)
            
            For Each xeRow In xeRows 'Loop Row Elements
                For Each xeCell In xeRow.ChildNodes 'Loop Column Elements
                    aryOutput(iRow, iCol) = xeCell.SelectSingleNode("ns:Data").nodeTypedValue
    '                MsgBox xeCell.SelectSingleNode("ns:Data").NodeType
    '                MsgBox xeCell.SelectSingleNode("ns:Data").nodeName
                    iCol = iCol + 1
                Next
                iCol = 0
                iRow = iRow + 1
                DoEvents
            Next
            ExcelXMLSpreadsheetToArray = aryOutput
        End If
    End If
    
errsub:
    Set xeCell = Nothing
    Set xeRow = Nothing
    Set xeRows = Nothing
    Set xmldoc = Nothing
End Function

Private Sub GetElementsByTagNameCodeSample()
''http://forum.videohelp.com/threads/245182-VBScript-to-read-XML-Assistance-Needed
'    Dim description, filepath
'
'    Set xmlDoc = CreateObject("Msxml2.DOMDocument")
'    xmlDoc.Load ("c:\test.xml")
'
'    Set ElemList = xmlDoc.getElementsByTagName("segment")
'    filepath = ElemList.Item(0).getAttribute("filePath")
'    MsgBox filepath
'
'    Set ElemList = xmlDoc.getElementsByTagName("description")
'    plot = ElemList.Item(0).Text
'    MsgBox plot
End Sub


'Sample code from http://objectmix.com/xml-soap/86701-xquery-call-vb6.html
Private Sub cmdChange_Click()
    Dim xmldoc As MSXML2.DOMDocument
    Dim xmlPath As String
    Dim xmlLoad As Boolean
    Dim att As String
    Dim xList As MSXML2.IXMLDOMNodeList
    Dim childNode As IXMLDOMNode
    
    Dim lbName As String
    Dim frmName As String
    
    xmlPath = "C:\Documents and Settings\User\Desktop\VB With XML\Xml File\Sample.xml"
    Set xmldoc = CreateObject("Microsoft.xmldom")
    xmlLoad = xmldoc.Load(xmlPath)

    If xmlLoad = True Then
        Set childNode = xmldoc.DocumentElement.SelectSingleNode("Language[@xml:lang='en']")
        
        'MsgBox xList(0).nodeName
        If Not (childNode Is Nothing) Then
            lbName = "lblName"
            frmName = "Customer"
            MsgBox childNode.SelectNodes("Form[@xml:frm='" & frmName & "']")(0).SelectNodes("Label[@xml:lbl='" & lbName & "']")(0).Text
        End If
        
        Set xList = xmldoc.DocumentElement.SelectNodes("Language/Form")
        MsgBox xList(0).Text
    End If
End Sub

Private Function IsXMLDocReady(xmldoc As Object, Optional TimeoutSeconds As Long = 30) As Boolean
    Dim dateCurrentTime As Date
    dateCurrentTime = Now
        
    Do While VBA.DateDiff("s", dateCurrentTime, Now) < TimeoutSeconds 'Check state
        If xmldoc.ReadyState > ReadyState.Loading Then
            IsXMLDocReady = True
            Exit Function
        End If
        VBA.Interaction.DoEvents
        Call Sleep(1000 / 4) 'milliseconds
    Loop
End Function

'=======================================================================================================
'Helper Functions
'=======================================================================================================
Private Function GetRegistryValueWMI(DataType As lpType, RegTreeKey As hKey, SubKeyName As String, ValueName As String) As Variant
'Registry Key Manipulation WMI Implementations
'http://msdn2.microsoft.com/en-us/library/aa394600.aspx
    On Error GoTo errsub

    Dim oReg As Object 'SWbemObjectEx
    Dim Value As Variant
    Dim bRet As Boolean
    Const COMPUTER As String = "."

    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & COMPUTER & "\root\default:StdRegProv")

    Select Case DataType
        Case lpType.REG_SZ, lpType.REG_EXPAND_SZ
            bRet = oReg.GetStringValue(RegTreeKey, SubKeyName, ValueName, Value)
            Value = CStr(Value)
        Case lpType.REG_DWORD
            bRet = oReg.GetDWORDValue(RegTreeKey, SubKeyName, ValueName, Value)
            Value = CLng(Value)
    End Select

    If bRet = False Then
        GetRegistryValueWMI = Value
    End If

errsub:
End Function
