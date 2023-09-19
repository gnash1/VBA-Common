VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'File:   clsCollection
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        2/13/12

'Description:
'Wrapper for Dictionary object used to avoid early binding. Also used as a base class for new classes.
'For early binding link to: "Microsoft Scripting Runtime"
'Added Items require both Key & Value

'Source:
'http://www.w3schools.com/asp/asp_ref_dictionary.asp
'http://visualbasic.about.com/od/usingvbnet/l/aa070903a.htm


'Default Member set in text mode to Item.
'http://www.cpearson.com/excel/DefaultMember.aspx
'Attribute Item.VB_UserMemId = 0

Public Enum CompareMethod
    BinaryCompare = 0   'vbBinaryCompare - binary comparison
    TextCompare = 1     'vbTextCompare - textual comparison
    DatabaseCompare = 2 'vbDatabaseCompare - database comparison
End Enum

Public Enum eSortBy
    DictionaryKeys = 0
    DictionaryItems = 1
End Enum

Private m_Dictionary As Object 'Dictionary

Private Sub Class_Initialize()
    Set m_Dictionary = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Set m_Dictionary = Nothing
End Sub

''''' Properties '''''
Public Property Get CompareMode() As CompareMethod
'Returns the comparison mode for comparing keys in a Dictionary object
    CompareMode = m_Dictionary.CompareMode
End Property

Public Property Let CompareMode(ByVal eCompareMethod As CompareMethod)
'Sets the comparison mode for comparing keys in a Dictionary object
    m_Dictionary.CompareMode = eCompareMethod
End Property

Public Property Get Count() As Long
'Returns the number of key/item pairs in a Dictionary object
    Count = m_Dictionary.Count
End Property

Public Property Get ItemByIndex(Index As Long) As Variant 'Could be object or value
'Returns the value of an item in a Dictionary object by 1 based index.
    If TypeOf m_Dictionary.Item(m_Dictionary.Keys()(Index - 1)) Is Object  Then
        Set ItemByIndex = m_Dictionary.Item(m_Dictionary.Keys()(Index - 1))
    Else
        ItemByIndex = m_Dictionary.Item(m_Dictionary.Keys()(Index - 1))
    End If
End Property

Public Property Get Item(Key As Variant) As Variant 'Could be object or value
'Returns the value of an item in a Dictionary object
'Set to default Item in text editor with "Attribute Value.VB_UserMemId = 0"
    If TypeOf m_Dictionary.Item(Key) Is Object  Then
        Set Item = m_Dictionary.Item(Key)
    Else
        Item = m_Dictionary.Item(Key)
    End If
End Property

Public Property Let Item(Key As Variant, Value As Variant)
'Sets the value of an item in a Dictionary object
    If TypeOf Value Is Object  Then
        Set m_Dictionary.Item(Key) = Value
    Else
        m_Dictionary.Item(Key) = Value
    End If
End Property

Public Property Get Key(Index As Long) As Variant 'Could be string or value
'Returns the value of a key in a Dictionary object
    Key = m_Dictionary.Keys()(Index)
End Property

''''' Methods '''''
Public Function Add(Key As Variant, Item As Variant) As Boolean
'Adds a new key/item pair to a Dictionary object
    On Error GoTo errsub
    Call m_Dictionary.Add(Key, Item)
    Add = True
errsub:
End Function

'Public Function Insert(Key As Variant, Item As Variant, Index As Long) As Boolean
'    On Error GoTo errsub
'    Dim i As Long
'    Dim NewDictionary As New Dictionary
'
'    Key = objDict.Key(i)
'    Item = objDict.ItemByIndex(i)
'
'    For i = 1 To m_Dictionary.Count
'        Call NewDictionary.Add(objDict.Key(i), objDict.ItemByIndex(i))
'    Next
'
''    Call m_Dictionary.Add(Key, Item)
''    Add = True
'errsub:
'End Function

Public Function Exists(Key As Variant) As Boolean
'Returns a Boolean value that indicates whether a specified key exists in the Dictionary object
    Exists = m_Dictionary.Exists(Key)
End Function

Public Function Items(Optional Transpose As Boolean = False) As Variant 'Array
'Returns an array of all the items in a Dictionary object.
    If Transpose = True Then
        Items = TransposeArray(m_Dictionary.Items())
    Else
        Items = m_Dictionary.Items()
    End If
End Function

Public Function Keys(Optional Transpose As Boolean = False) As Variant 'Array
'Returns an array of all the keys in a Dictionary object.
    If Transpose = True Then
        Keys = TransposeArray(m_Dictionary.Keys())
    Else
        Keys = m_Dictionary.Keys()
    End If
End Function

Public Function Remove(Key As Variant) As Boolean
'Removes one specified key/item pair from the Dictionary object
    On Error GoTo errsub
    m_Dictionary.Remove (Key)
errsub:
End Function

Public Function Clear()
'Public Function RemoveAll()
'Removes all the key/item pairs in the Dictionary object
    Call m_Dictionary.RemoveAll
End Function

'Public Sub SortDictionary(objDict As clsDictionary, SortBy As eSortBy)
''http://support.microsoft.com/kb/246067
'    Dim strDict()
'    Dim objKey
'    Dim strKey, strItem
'    Dim X, Y
'    Dim i As Long
'
'    If objDict.Count > 1 Then
'        ' create an array to store dictionary information
'        ReDim strDict(objDict.Count, 2)
'
'        ' populate the string array
''        For i = 1 To objDict.Count
''        For Each objKey In objDict.Count
'            strDict(X, dictKey) = CStr(objKey)
'            strDict(X, dictItem) = CStr(objDict(objKey))
'            X = X + 1
'        Next
'
'        ' perform a a shell sort of the string array
'        For X = 0 To (objDict.Count - 2)
'            For Y = X To (objDict.Count - 1)
'                If StrComp(strDict(X, SortBy), strDict(Y, SortBy), vbTextCompare) > 0 Then
'                    strKey = strDict(X, dictKey)
'                    strItem = strDict(X, dictItem)
'                    strDict(X, dictKey) = strDict(Y, dictKey)
'                    strDict(X, dictItem) = strDict(Y, dictItem)
'                    strDict(Y, dictKey) = strKey
'                    strDict(Y, dictItem) = strItem
'                End If
'            Next
'        Next
'
'        ' erase the contents of the dictionary object
'        objDict.RemoveAll
'
'      ' repopulate the dictionary with the sorted information
'        For X = 0 To (objDict.Count - 1)
'            objDict.Add strDict(X, dictKey), strDict(X, dictItem)
'        Next
'    End If
'End Sub

'Private Sub Sort_STRING(ByRef strArray() As String, Optional ByVal blnDecending As Boolean = False)
'   Dim strTempVal  As String
'   Dim lngCounter  As Long
'   Dim lngGapSize  As Long
'   Dim lngCurPos   As Long
'   Dim lngFirstRow As Long
'   Dim lngLastRow  As Long
'   Dim lngNumRows  As Long
'
'   lngFirstRow = LBound(strArray)
'   lngLastRow = UBound(strArray)
'   lngNumRows = lngLastRow - lngFirstRow + 1
'
'   Do
'      lngGapSize = lngGapSize * 3 + 1
'   Loop Until lngGapSize > lngNumRows
'
'   Do
'      lngGapSize = lngGapSize \ 3
'      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
'         lngCurPos = lngCounter
'         strTempVal = strArray(lngCounter)
'         Do While CompareResult_TXT(strArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
'            strArray(lngCurPos) = strArray(lngCurPos - lngGapSize)
'            lngCurPos = lngCurPos - lngGapSize
'            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
'         Loop
'         strArray(lngCurPos) = strTempVal
'      Next
'   Loop Until lngGapSize = 1
'End Sub

'Public Function SortCollection(ByVal c As Collection) As Collection
''http://www.source-code.biz/snippets/vbasic/6.htm
'' This routine uses the "heap sort" algorithm to sort a VB collection.
'' It returns the sorted collection.
'' Author: Christian d'Heureuse (www.source-code.biz)
'    Dim n As Long
'    n = c.Count
'
'    If n = 0 Then
'        Set SortCollection = New Collection
'        Exit Function
'    End If
'
'    ' allocate index array
'    ReDim Index(0 To n - 1) As Long
'    Dim i As Long, m As Long
'
'    ' fill index array
'    For i = 0 To n - 1
'        Index(i) = i + 1
'    Next
'
'    ' generate ordered heap
'    For i = n \ 2 - 1 To 0 Step -1
'        Heapify c, Index, i, n
'    Next
'
'    ' sort the index array
'    For m = n To 2 Step -1
'        Exchange Index, 0, m - 1    ' move highest element to top
'        Heapify c, Index, 0, m - 1
'    Next
'
'    ' fill output collection
'    Dim c2 As New Collection
'    For i = 0 To n - 1
'        c2.Add c.Item(Index(i))
'    Next
'
'    Set SortCollection = c2
'End Function

'============================================================================================================
'==Helper Functions
'============================================================================================================
Private Function TransposeArray(vArray As Variant) As Variant
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
