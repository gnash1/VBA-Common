Attribute VB_Name = "modSort"
Option Explicit
'Multiple Sort methods from various sources.

'Modified from http://support.microsoft.com/kb/q133135
'Option Base 1

'Bubble Sort (fix due to removal of Option Base 1)
Function BubbleSort(TempArray As Variant)
    Dim Temp As Variant
    Dim i As Long
    Dim bExchanges As Boolean 'Long
    ' Loop until no "exchanges" are made.
    Do
        bExchanges = False
        ' Loop through each element in the array.
        For i = 1 To UBound(TempArray) - 1
            'If the element > the element following it, exchange the two elements.
            If TempArray(i) > TempArray(i + 1) Then
                bExchanges = True
                Temp = TempArray(i)
                TempArray(i) = TempArray(i + 1)
                TempArray(i + 1) = Temp
            End If
        Next i
    Loop While bExchanges
End Function

'Selection Sort(fix due to removal of Option Base 1)
Function SelectionSort(TempArray As Variant)
    Dim MaxVal As Variant
    Dim MaxIndex As Long
    Dim i, j As Long
    ' Step through the elements in the array starting with the last element in the array.
    For i = UBound(TempArray) To 1 Step -1
        ' Set MaxVal to the element in the array and save the index of this element as MaxIndex.
        MaxVal = TempArray(i)
        MaxIndex = i

        ' Loop through the remaining elements to see if any is
        ' larger than MaxVal. If it is then set this element to be the new MaxVal.
        For j = 1 To i
            If TempArray(j) > MaxVal Then
                MaxVal = TempArray(j)
                MaxIndex = j
            End If
        Next j

        ' If the index of the largest element is not i, then exchange this element with element i.
        If MaxIndex < i Then
            TempArray(MaxIndex) = TempArray(i)
            TempArray(i) = MaxVal
        End If
    Next i
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
Public Function SortCollection(ByVal c As Collection) As Collection
'http://www.source-code.biz/snippets/vbasic/6.htm
' This routine uses the "heap sort" algorithm to sort a VB collection.
' It returns the sorted collection.
' Author: Christian d'Heureuse (www.source-code.biz)
    Dim n As Long
    n = c.count
    
    If n = 0 Then
        Set SortCollection = New Collection
        Exit Function
    End If
    
    ' allocate index array
    ReDim Index(0 To n - 1) As Long
    Dim i As Long, m As Long
    
    ' fill index array
    For i = 0 To n - 1
        Index(i) = i + 1
    Next
    
    ' generate ordered heap
    For i = n \ 2 - 1 To 0 Step -1
        Heapify c, Index, i, n
    Next
    
    ' sort the index array
    For m = n To 2 Step -1
        Exchange Index, 0, m - 1    ' move highest element to top
        Heapify c, Index, 0, m - 1
    Next
    
    ' fill output collection
    Dim c2 As New Collection
    For i = 0 To n - 1
        c2.Add c.Item(Index(i))
    Next
    
    Set SortCollection = c2
End Function

Private Sub Heapify(ByVal c As Collection, Index() As Long, ByVal i1 As Long, ByVal n As Long)
    ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
    Dim nDiv2 As Long
    Dim i As Long
    Dim k As Long
    
    nDiv2 = n \ 2
    i = i1
    
    Do While i < nDiv2
        k = 2 * i + 1
        If k + 1 < n Then
            If LCase(c.Item(Index(k))) < LCase(c.Item(Index(k + 1))) Then
                k = k + 1
            End If
        End If
        
        If c.Item(Index(i)) >= c.Item(Index(k)) Then
            Exit Do
        End If
        
        Exchange Index, i, k
        i = k
    Loop
End Sub

Private Sub Exchange(Index() As Long, ByVal i As Long, ByVal j As Long)
    Dim Temp As Long
    Temp = Index(i)
    Index(i) = Index(j)
    Index(j) = Temp
End Sub

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

'http://www.freevbcode.com/ShowCode.asp?ID=3645
Public Sub SortCollectionSimple(ColVar As Collection)
    Dim oCol As Collection
    Dim i As Integer
    Dim i2 As Integer
    Dim iBefore As Integer
    If Not (ColVar Is Nothing) Then
        If ColVar.count > 0 Then
            Set oCol = New Collection
            For i = 1 To ColVar.count
                If oCol.count = 0 Then
                    oCol.Add ColVar(i)
                Else
                    iBefore = 0
                    For i2 = oCol.count To 1 Step -1
                        If LCase(ColVar(i)) < LCase(oCol(i2)) Then
                            iBefore = i2
                        Else
                            Exit For
                        End If
                    Next
                    If iBefore = 0 Then
                        oCol.Add ColVar(i)
                    Else
                        oCol.Add ColVar(i), , iBefore
                    End If
                End If
            Next
            Set ColVar = oCol
            Set oCol = Nothing
        End If
    End If
End Sub


'
' modSort Module
' --------------
'
' Created By  : Kevin Wilson
'               http://www.TheVBZone.com   ( The VB Zone )
'               http://www.TheVBZone.net   ( The VB Zone .net )
'
' Created On  : June 23, 2003
' Last Update : June 23, 2003
'
' VB Versions : 5.0 / 6.0
'
' Requires    : NOTHING
'
' Description : This module is designed to make it easy to sort data arrays of different data types.  The
'               sort functions used in this module use an sorting algorithm called the "Shell Sort".  The
'               "Shell Sort" was invented by Donald Shell in 1959.  It is much more efficient than other
'               such O(n2) sort algorithms as the "Bubble Sort", "Selection Sort", and "Insertion Sort"
'               but still is fairly simplistic in it's design.  There other O(n log n) sorting methods
'               that are more efficient such as the "Heap Sort", "Merge Sort", and "Quick Sort"...
'               but these sorting algorithms are MUCH more complex and memory intensive.  They are also
'               oriented towards sorting numbers, not sorting strings.
'
' NOTE        : In my testing... I sorted both a string array and a number array containing 100,000 elements
'               each and both string arrays, and number arrays were sorted correctly within 5 to 10 seconds.
'
' See Also    : http://linux.wku.edu/~lamonml/algor/sort/index.html
'               http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnoffpro01/html/ABetterShellSortPartI.asp
'
' Example Use :
'
'   Option Explicit
'   Private Sub Form_Load()
'      Dim MyArray(4) As String
'      Dim bytCounter As Byte
'      MyArray(0) = "Hello"
'      MyArray(1) = "My"
'      MyArray(2) = "Name"
'      MyArray(3) = "Is"
'      MyArray(4) = "Kevin"
'      Call Sort_STRING(MyArray, True)
'      Me.AutoRedraw = True
'      For bytCounter = 0 To 4
'         Me.Print MyArray(bytCounter)
'      Next
'   End Sub
'
'=============================================================================================================
'
' LEGAL:
'
' You are free to use this code as long as you keep the above heading information intact and unchanged. Credit
' given where credit is due.  Also, it is not required, but it would be appreciated if you would mention
' somewhere in your compiled program that that your program makes use of code written and distributed by
' Kevin Wilson (www.TheVBZone.com).  Feel free to link to this code via your web site or articles.
'
' You may NOT take this code and pass it off as your own.  You may NOT distribute this code on your own server
' or web site.  You may NOT take code created by Kevin Wilson (www.TheVBZone.com) and use it to create products,
' utilities, or applications that directly compete with products, utilities, and applications created by Kevin
' Wilson, TheVBZone.com, or Wilson Media.  You may NOT take this code and sell it for profit without first
' obtaining the written consent of the author Kevin Wilson.
'
' These conditions are subject to change at the discretion of the owner Kevin Wilson at any time without
' warning or notice.  Copyright© by Kevin Wilson.  All rights reserved.
'
'=============================================================================================================




'=============================================================================================================
' Sort_* Functions
'
' The Sort_BYTE, Sort_INTEGER, Sort_LONG, Sort_SINGLE, Sort_DOUBLE, Sort_DATE, and Sort_STRING functions take
' in the respective data type arrays and sort them either ascending or decending depending on the parameters
' passed to the functions.
'
' Parameter:                 Use:
' -------------------------------------------
' strArray()                 References the array to be sorted.
' blnDecending               If FALSE (default), the array will be sorted in ascending order (A,B,C).  If
'                            TRUE, the array will be sorted in descending order (Z,X,Y).
'
' Return:
' -------
' (none)
'
'=============================================================================================================
Public Sub Sort_STRING(ByRef strArray() As String, Optional ByVal blnDecending As Boolean = False)
   
   Dim strTempVal  As String
   Dim lngCounter  As Long
   Dim lngGapSize  As Long
   Dim lngCurPos   As Long
   Dim lngFirstRow As Long
   Dim lngLastRow  As Long
   Dim lngNumRows  As Long
   
   lngFirstRow = LBound(strArray)
   lngLastRow = UBound(strArray)
   lngNumRows = lngLastRow - lngFirstRow + 1
   
   Do
      lngGapSize = lngGapSize * 3 + 1
   Loop Until lngGapSize > lngNumRows
   
   Do
      lngGapSize = lngGapSize \ 3
      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
         lngCurPos = lngCounter
         strTempVal = strArray(lngCounter)
         Do While CompareResult_TXT(strArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
            strArray(lngCurPos) = strArray(lngCurPos - lngGapSize)
            lngCurPos = lngCurPos - lngGapSize
            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
         Loop
         strArray(lngCurPos) = strTempVal
      Next
   Loop Until lngGapSize = 1
   
End Sub

' (See documentation for the "Sort_STRING" function)
Public Sub Sort_BYTE(ByRef bytArray() As Byte, Optional ByVal blnDecending As Boolean = False)
   
   Dim strTempVal  As String
   Dim lngCounter  As Long
   Dim lngGapSize  As Long
   Dim lngCurPos   As Long
   Dim lngFirstRow As Long
   Dim lngLastRow  As Long
   Dim lngNumRows  As Long
   
   lngFirstRow = LBound(bytArray)
   lngLastRow = UBound(bytArray)
   lngNumRows = lngLastRow - lngFirstRow + 1
   
   Do
      lngGapSize = lngGapSize * 3 + 1
   Loop Until lngGapSize > lngNumRows
   
   Do
      lngGapSize = lngGapSize \ 3
      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
         lngCurPos = lngCounter
         strTempVal = bytArray(lngCounter)
         Do While CompareResult_NUM(bytArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
            bytArray(lngCurPos) = bytArray(lngCurPos - lngGapSize)
            lngCurPos = lngCurPos - lngGapSize
            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
         Loop
         bytArray(lngCurPos) = strTempVal
      Next
   Loop Until lngGapSize = 1
   
End Sub

' (See documentation for the "Sort_STRING" function)
Public Sub Sort_INTEGER(ByRef intArray() As Integer, Optional ByVal blnDecending As Boolean = False)
   
   Dim strTempVal  As String
   Dim lngCounter  As Long
   Dim lngGapSize  As Long
   Dim lngCurPos   As Long
   Dim lngFirstRow As Long
   Dim lngLastRow  As Long
   Dim lngNumRows  As Long
   
   lngFirstRow = LBound(intArray)
   lngLastRow = UBound(intArray)
   lngNumRows = lngLastRow - lngFirstRow + 1
   
   Do
      lngGapSize = lngGapSize * 3 + 1
   Loop Until lngGapSize > lngNumRows
   
   Do
      lngGapSize = lngGapSize \ 3
      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
         lngCurPos = lngCounter
         strTempVal = intArray(lngCounter)
         Do While CompareResult_NUM(intArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
            intArray(lngCurPos) = intArray(lngCurPos - lngGapSize)
            lngCurPos = lngCurPos - lngGapSize
            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
         Loop
         intArray(lngCurPos) = strTempVal
      Next
   Loop Until lngGapSize = 1
   
End Sub

' (See documentation for the "Sort_STRING" function)
Public Sub Sort_LONG(ByRef lngArray() As Long, Optional ByVal blnDecending As Boolean = False)
   
   Dim strTempVal  As String
   Dim lngCounter  As Long
   Dim lngGapSize  As Long
   Dim lngCurPos   As Long
   Dim lngFirstRow As Long
   Dim lngLastRow  As Long
   Dim lngNumRows  As Long
   
   lngFirstRow = LBound(lngArray)
   lngLastRow = UBound(lngArray)
   lngNumRows = lngLastRow - lngFirstRow + 1
   
   Do
      lngGapSize = lngGapSize * 3 + 1
   Loop Until lngGapSize > lngNumRows
   
   Do
      lngGapSize = lngGapSize \ 3
      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
         lngCurPos = lngCounter
         strTempVal = lngArray(lngCounter)
         Do While CompareResult_NUM(lngArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
            lngArray(lngCurPos) = lngArray(lngCurPos - lngGapSize)
            lngCurPos = lngCurPos - lngGapSize
            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
         Loop
         lngArray(lngCurPos) = strTempVal
      Next
   Loop Until lngGapSize = 1
   
End Sub

' (See documentation for the "Sort_STRING" function)
Public Sub Sort_SINGLE(ByRef sngArray() As Single, Optional ByVal blnDecending As Boolean = False)
   
   Dim strTempVal  As String
   Dim lngCounter  As Long
   Dim lngGapSize  As Long
   Dim lngCurPos   As Long
   Dim lngFirstRow As Long
   Dim lngLastRow  As Long
   Dim lngNumRows  As Long
   
   lngFirstRow = LBound(sngArray)
   lngLastRow = UBound(sngArray)
   lngNumRows = lngLastRow - lngFirstRow + 1
   
   Do
      lngGapSize = lngGapSize * 3 + 1
   Loop Until lngGapSize > lngNumRows
   
   Do
      lngGapSize = lngGapSize \ 3
      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
         lngCurPos = lngCounter
         strTempVal = sngArray(lngCounter)
         Do While CompareResult_NUM(sngArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
            sngArray(lngCurPos) = sngArray(lngCurPos - lngGapSize)
            lngCurPos = lngCurPos - lngGapSize
            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
         Loop
         sngArray(lngCurPos) = strTempVal
      Next
   Loop Until lngGapSize = 1
   
End Sub

' (See documentation for the "Sort_STRING" function)
Public Sub Sort_DOUBLE(ByRef dblArray() As Double, Optional ByVal blnDecending As Boolean = False)
   
   Dim strTempVal  As String
   Dim lngCounter  As Long
   Dim lngGapSize  As Long
   Dim lngCurPos   As Long
   Dim lngFirstRow As Long
   Dim lngLastRow  As Long
   Dim lngNumRows  As Long
   
   lngFirstRow = LBound(dblArray)
   lngLastRow = UBound(dblArray)
   lngNumRows = lngLastRow - lngFirstRow + 1
   
   Do
      lngGapSize = lngGapSize * 3 + 1
   Loop Until lngGapSize > lngNumRows
   
   Do
      lngGapSize = lngGapSize \ 3
      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
         lngCurPos = lngCounter
         strTempVal = dblArray(lngCounter)
         Do While CompareResult_NUM(dblArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
            dblArray(lngCurPos) = dblArray(lngCurPos - lngGapSize)
            lngCurPos = lngCurPos - lngGapSize
            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
         Loop
         dblArray(lngCurPos) = strTempVal
      Next
   Loop Until lngGapSize = 1
   
End Sub

' (See documentation for the "Sort_STRING" function)
Public Sub Sort_DATE(ByRef datArray() As Double, Optional ByVal blnDecending As Boolean = False)
   
   Dim strTempVal  As String
   Dim lngCounter  As Long
   Dim lngGapSize  As Long
   Dim lngCurPos   As Long
   Dim lngFirstRow As Long
   Dim lngLastRow  As Long
   Dim lngNumRows  As Long
   
   lngFirstRow = LBound(datArray)
   lngLastRow = UBound(datArray)
   lngNumRows = lngLastRow - lngFirstRow + 1
   
   Do
      lngGapSize = lngGapSize * 3 + 1
   Loop Until lngGapSize > lngNumRows
   
   Do
      lngGapSize = lngGapSize \ 3
      For lngCounter = (lngGapSize + lngFirstRow) To lngLastRow
         lngCurPos = lngCounter
         strTempVal = datArray(lngCounter)
         Do While CompareResult_DAT(datArray(lngCurPos - lngGapSize), strTempVal, blnDecending)
            datArray(lngCurPos) = datArray(lngCurPos - lngGapSize)
            lngCurPos = lngCurPos - lngGapSize
            If (lngCurPos - lngGapSize) < lngFirstRow Then Exit Do
         Loop
         datArray(lngCurPos) = strTempVal
      Next
   Loop Until lngGapSize = 1
   
End Sub

' This function is used within this module only to compare values
Private Function CompareResult_TXT(ByVal strValue1 As String, ByVal strValue2 As String, Optional blnDescending As Boolean = False) As Boolean
   CompareResult_TXT = CBool(StrComp(strValue1, strValue2, vbTextCompare) = 1)
   CompareResult_TXT = CompareResult_TXT Xor blnDescending
End Function

' This function is used within this module only to compare values
Private Function CompareResult_NUM(ByVal dblValue1 As Double, ByVal dblValue2 As Double, Optional blnDescending As Boolean = False) As Boolean
   CompareResult_NUM = CBool(dblValue1 > dblValue2)
   CompareResult_NUM = CompareResult_NUM Xor blnDescending
End Function

' This function is used within this module only to compare values
Private Function CompareResult_DAT(ByVal datValue1 As Date, ByVal datValue2 As Date, Optional blnDescending As Boolean = False) As Boolean
   CompareResult_DAT = CBool(datValue1 > datValue2)
   CompareResult_DAT = CompareResult_DAT Xor blnDescending
End Function

'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'Title: Sort String Arrays without Swapping Strings
'Description: Sort Array without Swapping Strings
'The latest Sort submission by Ulli, whcih is unquestionably superior and advanced is, well, advanced! so i cooked-up this sort which is very simple, the algorithm used is shell sort but instead of swapping strings, we swap their index, and finally return an array of sorted indexes, well sort of (pun intended ;-)
'This file came from Planet-Source-Code.com...the home millions of lines of source code
'You can view comments on this code/and or vote on it at: http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=57898&lngWId=1

Public Function ShellSortLong(SortArray() As String, Optional ByVal IgnoreCase As Boolean = True) As Long()
Dim sVal1 As String, sVal2 As String
Dim IndexArray() As Long
Dim idx As Long, Row As Long, MaxRow As Long, MinRow As Long
Dim Swtch As Long, Limit As Long, Offset As Long

MaxRow = UBound(SortArray)
MinRow = LBound(SortArray)
ReDim IndexArray(MinRow To MaxRow)
For idx = MinRow To MaxRow
    IndexArray(idx) = idx
Next

Offset = MaxRow \ 2

Do While Offset > 0
      Limit = MaxRow - Offset
      Do
         Swtch = False         'Assume no switches at this offset.

         ' Compare elements and switch ones out of order:
         For Row = MinRow To Limit
                sVal1 = SortArray(IndexArray(Row))
                sVal2 = SortArray(IndexArray(Row + Offset))
                If IgnoreCase Then
                    sVal1 = LCase(sVal1)
                    sVal2 = LCase(sVal2)
                End If
                If sVal1 > sVal2 Then
                   SwapLongs IndexArray(Row), IndexArray(Row + Offset)
                   Swtch = Row
                End If
         Next Row

         ' Sort on next pass only to where last switch was made:
         Limit = Swtch - Offset
      Loop While Swtch

      ' No switches at last offset, try one half as big:
      Offset = Offset \ 2
   Loop

ShellSortLong = IndexArray

End Function
Private Sub SwapLongs(ByRef var1 As Long, ByRef var2 As Long)

    Dim X As Long
    X = var1
    var1 = var2
    var2 = X

End Sub
