VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCollection"
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
'Wrapper for Collection object to add additional functionality similar to Dictionay object. Also used as a base class for new classes.
'Added Items don't require both Key & Value, as Dictionary object does.

'Can set default member (in text mode) to Item.
'http://www.cpearson.com/excel/DefaultMember.aspx

'Public Enum CompareMethod
'    BinaryCompare = 0   'vbBinaryCompare - binary comparison
'    TextCompare = 1     'vbTextCompare - textual comparison
'    DatabaseCompare = 2 'vbDatabaseCompare - database comparison
'End Enum

Private m_Collection As Collection
Private m_Keys() As Variant 'Could add to keep track of keys (like dictionary object).

Private Sub Class_Initialize()
    Set m_Collection = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_Collection = Nothing
End Sub

''''' Properties '''''
'Public Property Get CompareMode() As CompareMethod
''Returns the comparison mode for comparing keys in a Dictionary object
'    CompareMode = m_Collection.CompareMode
'End Property

'Public Property Let CompareMode(ByVal eCompareMethod As CompareMethod)
''Sets the comparison mode for comparing keys in a Dictionary object
'    m_Collection.CompareMode = eCompareMethod
'End Property

Public Property Get Count() As Long
'Returns the number of key/item pairs in a Dictionary object
    Count = m_Collection.Count
End Property

Public Property Get Item(Key As Variant) As Object
'Returns the value of an item in a Dictionary object
'Key value must start with a character
     Set Item = m_Collection(Key)
End Property

Public Property Let Item(Key As Variant, Value As Variant)
'Sets the value of an item in a Dictionary object
'    If Exists(Key) Then 'Check this before calling.
    m_Collection(Key) = Value
'    End If
End Property

''''' Methods '''''
Public Function Add(Item As Variant, Optional Key As Variant) ', Optional Before As Variant, Optional After As Variant) As Boolean
'Adds a new key/item pair to a Dictionary object
    On Error GoTo errsub
    If IsMissing(Key) = False Then
        Call m_Collection.Add(Item, Key)
    Else
        Call m_Collection.Add(Item)
    End If
    
    Add = True
errsub:
End Function

Public Function Exists(Key As Variant) As Boolean
'Returns a Boolean value that indicates whether a specified key exists in the Collection object
    On Error GoTo errsub
    Dim varTemp As Variant
    
    Set varTemp = m_Collection.Item(Key)
    Exists = Not (Err.Number = 5 Or Err.Number = 9)
errsub:
End Function

Public Function Items() As Variant 'Array
'Returns an array of all the items in a Collection.
    Dim lCount As Long
    Dim vArray As Variant
    Dim vItem As Variant

    ReDim vArray(m_Collection.Count - 1) As Variant

    For Each vItem In m_Collection
        vArray(lCount) = vItem
        lCount = lCount + 1
    Next

    Items = vArray
End Function

'Public Function Keys() As Variant 'Array
''Returns an array of all the keys in a Dictionary object.
'    Keys = m_Collection.Keys()
'End Function

Public Function Remove(Key As Variant) As Boolean
'Removes one specified key/item pair from the Dictionary object
    On Error GoTo errsub
    m_Collection.Remove (Key)
    Remove = True
errsub:
End Function

Public Function Clear()
'Public Function RemoveAll()
'Removes all the key/item pairs in the Dictionary object
    Do While m_Collection.Count > 0
        Call m_Collection.Remove(1)
    Loop
End Function

Public Function NewEnum() As IUnknown
'Enable For...Each loops 'Also requires hidden property to be put in via text editor
'Attribute NewEnum.VB_UserMemID = -4
    Set NewEnum = m_Collection.[_NewEnum]
End Function

'SortCollection

'Private Function RemoveCollectionDuplicates(ByVal cOriginal As Collection, cNew As Collection) As Collection
''Remove duplicate New collection items from Original collection
'    Dim cItemOrig As Variant
'    Dim cItemNew As Variant
'    Dim i As Long
'
'    If Not cOriginal Is Nothing And Not cNew Is Nothing Then
''        Set RemoveCollectionDuplicates = cItemNew
'        For Each cItemNew In cNew
'            i = 1
'            For Each cItemOrig In cOriginal
'                If StrComp(cItemOrig, cItemNew) = 0 Then 'Same
'                    Call cOriginal.Remove(i)
'                    Exit For
'                End If
'                i = i + 1
'            Next
'        Next
'
'        If cOriginal.Count > 0 Then
'            Set RemoveCollectionDuplicates = cOriginal
'        End If
'    ElseIf cOriginal Is Nothing And Not cNew Is Nothing Then
'        If cNew.Count > 0 Then
'            Set RemoveCollectionDuplicates = cNew
'        End If
'    End If
'End Function

Public Function CollectionToArray(col As Collection) As Variant
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
