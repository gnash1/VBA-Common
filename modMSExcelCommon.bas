Attribute VB_Name = "modMSExcelCommon"
'File:   modMSExcelCommon
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:       3/25/13

'To do:
'Remove references to "Range" .Cells.Count, which are not always the last cell (right/bottom) in range

'Summary:
'Excel common code used in multiple applications.
'Assumes that use will be within Excel (with early bound reference)
'If used outside Excel late bound references should be substituted.
'ProModel Corporation - www.promodel.com
'Contains late binding
'Note code is organized into related sections.
'All code is prefered to be generic and designed to function regardless of focus in Excel.
'Don't reply on ActiveWorkbook/Workbook, ActiveSheet/Worksheet or ActiveCell or call Select
'
'Online References:
'http://www.cpearson.com/excel
'http://www.contextures.com
'http://www.j-walk.com/ss/excel/tips/

'Revisions:
'Date     Initials    Description of changes

'General Notes:
'See Excel.Constants for application constants.
'Application.Volatile True - to cause autoupdate when values are modified.
'Range(cells.count) - Gives offset from starting cell, which is equal to last cell, but only if range is contiguous.  If not contiguous have to get offset in the last area.

'Warnings:
'Any function that uses xlUp won't look at invisible rows/columns.
'Used range will only provide Used Range and may not be what is intended when offsetting as UsedRange doesn't start at A1.

'Any
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''ToC:
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Pivot Table Functions
'' Excel Auto Filters
'' Chart Functions
'' Range functions
'' Workbook & Worksheet functions
'' Named Ranges/Formulas
'' File/Directory/Path Functions
'' Control/Object/Collection Functions
'' General Utility Functions
'' Custom Property Functions
'' Outline functions
'' Validation functions
'' Graphic/Picture Functions
'' Protect & Unprotect Functions
'' Unsorted'

Option Explicit

'Supporting Enum should be Private
Private Enum PathParseMode
    Path
    FileName
    FileExtension
    FileNameWithoutExtension
End Enum

Private Enum RelativePosition
    Before
    After
End Enum

Private Enum Orientation
    Horizontal
    Vertical
End Enum

Private Enum cellProperty
    Value
'    Background_Color
'    Text_Color
    'Add others
End Enum

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

'May conflict with RECT defined in modAPI
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Const WM_USER = 1024
Private Const MAX_PATH = 260
Private Const UNIQUE_NAME = &H0

Private Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long
Private Declare Function CoCreateGuid Lib "ole32.dll" (rclsid As GUID) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetTempFileNameA Lib "kernel32" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Dim Test As Variant

'For status bar
'Dim m_frmHwnd As Long
'Dim m_pbHwnd As Long
'Dim hwndStatusbar As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Pivot Table Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'PivotTable Ranges
'See formatting examples of ranges here: http://peltiertech.com/WordPress/referencing-pivot-table-ranges-in-vba/
'MsgBox .DataBodyRange.Address 'C3:R32
'MsgBox .DataLabelRange.Address 'A1
'MsgBox .ColumnRange.Address 'C1:R2
'MsgBox .RowRange.Address 'A2:B32
'MsgBox .TableRange1.Address 'A1:R32
'MsgBox .TableRange2.Address 'A1:R32

Private Sub CreateNewPivotTables(bDeletePreviousTables As Boolean, wsActiveWorkSheet As Worksheet, rngSourceRange As Excel.Range, rngPivotTableAnchorCell As Excel.Range, strPivotTableName As String, _
    intColumnFieldIndex As Long, intRowFieldIndex As Long, intDataFieldIndex As Long, Optional intPageFieldIndex As Long, Optional intHiddenFieldIndex As Long)
'Pivot Tables: Create new Pivot Tables optionally deleteing old ones.
    'FieldIndex values must be > 0.  If 0 they are not included.
    Dim PTCache As PivotCache
    Dim PT As PivotTable

    If bDeletePreviousTables = True Then
        Call DeleteAllPivotTables(wsActiveWorkSheet)
    End If

    Set PTCache = wsActiveWorkSheet.Parent.PivotCaches.Add(xlDatabase, rngSourceRange)
    'Set PTCache = ActiveWorkbook.PivotCaches.Add(xlDatabase, rngSourceRange)
    Set PT = PTCache.CreatePivotTable(rngPivotTableAnchorCell, strPivotTableName)

    With PT
        'Index of mapping from selection column header index to Pivot Table fields.
        If (intColumnFieldIndex > 0) Then
            .PivotFields(intColumnFieldIndex).Orientation = xlColumnField
        End If

        If (intRowFieldIndex > 0) Then
            .PivotFields(intRowFieldIndex).Orientation = xlRowField
        End If

        If (intDataFieldIndex > 0) Then
            .PivotFields(intDataFieldIndex).Orientation = xlDataField
        End If

        If (intHiddenFieldIndex > 0) Then
            .PivotFields(intHiddenFieldIndex).Orientation = xlHidden
        End If

        If (intPageFieldIndex > 0) Then
            .PivotFields(intPageFieldIndex).Orientation = xlPageField
        End If
    End With
    'ActiveWorkbook.PivotCaches(1).RefreshOnFileOpen = True 'Also set in right click menu
    'ActiveWorkbook.PivotCaches.Item(1).Refresh 'Refresh cache
End Sub

Private Sub DeleteAllPivotTables(wsActiveWorkSheet As Worksheet)
'Pivot Tables: Clear out previous Pivot Tables
    Dim PT As PivotTable

    For Each PT In wsActiveWorkSheet.PivotTables
        PT.TableRange2.Clear
    Next PT
End Sub

Private Sub DeleteMissingItems2002All(wsActiveWorkSheet As Worksheet)
'Pivot Tables: In Excel 2002, and later versions, you can programmatically change the pivot table properties, to prevent missing items from appearing, or clear items that have appeared.
'prevents unused items in non-OLAP PivotTables

    'in Excel 2002 and later versions
    'If unused items already exist,
      'run this macro then refresh the table
    Dim PT As PivotTable
    Dim ws As Worksheet

    For Each ws In wsActiveWorkSheet.Parent.Worksheets
    'For Each ws In ActiveWorkbook.Worksheets
      For Each PT In ws.PivotTables
        PT.PivotCache.MissingItemsLimit = xlMissingItemsNone
        PT.PivotCache.Refresh
      Next PT
    Next ws
End Sub

Private Sub DeleteOldItemsWB(wsActiveWorkSheet As Worksheet)
'Pivot Tables: In previous versions of Excel, run the following code to clear the old items from the dropdown list of a Pivot Table.
    'gets rid of unused items in PivotTable
    ' based on MSKB (202232)
    Dim ws As Worksheet
    Dim PT As PivotTable
    Dim pf As PivotField
    Dim pi As PivotItem

    On Error Resume Next
    For Each ws In wsActiveWorkSheet.Parent.Worksheets
    'For Each ws In ActiveWorkbook.Worksheets
      For Each PT In ws.PivotTables
        PT.RefreshTable
        PT.ManualUpdate = True
        For Each pf In PT.VisibleFields
          If pf.name <> "Data" Then
            For Each pi In pf.PivotItems
              If pi.RecordCount = 0 And _
                Not pi.IsCalculated Then
                pi.Delete
              End If
            Next pi
          End If
        Next pf
        PT.ManualUpdate = False
        PT.RefreshTable
      Next PT
    Next ws
End Sub

Private Function GetPivotTableTotalsRows(PT As Excel.PivotTable, bIncludeGrandTotal As Boolean) As Collection
'Returns a collection of ranges containing SubTotal Rows
'http://peltiertech.com/WordPress/referencing-pivot-table-ranges-in-vba/
    Dim colResults As Collection
    Dim rngAreas As Excel.Range
    Dim pf As Excel.PivotField
    Dim pi As Excel.PivotItem
    Dim rngArea As Excel.Range

    On Error Resume Next

    For Each pf In PT.PivotFields
        If pf.Orientation = xlRowField And pf.Subtotals(1) = True Then
            For Each pi In pf.PivotItems
                Set rngAreas = AddRange(rngAreas, pi.LabelRange)
            Next
        End If
    Next

    If Not rngAreas Is Nothing Then
        'Need to consider multiple sub total levels? pf.TotalLevels
        Set colResults = New Collection
        For Each rngArea In rngAreas.Areas
            colResults.Add (Application.Intersect(PT.TableRange1, rngArea(rngArea.Cells.Count).Offset(1).EntireRow))
        Next
    End If

    If bIncludeGrandTotal = True And PT.ColumnGrand = True Then
        colResults.Add Application.Intersect(PT.TableRange1, PT.TableRange1(PT.TableRange1.Cells.Count).EntireRow)
    End If

    Set GetPivotTableTotalsRows = colResults

End Function

Private Function TogglePTSubtotalsOff(PT As Excel.PivotTable) As Boolean
'Untested
'Turn off subtotals in Pivot Tables.
    On Error GoTo errsub
    Dim ptField As Excel.PivotField
    Dim vSubTotal As Variant

'    PT.SubtotalLocation xlAtBottom 'SubTotals as Bottom
'    PT.SubtotalLocation xlAtTop 'SubTotals as Top

    For Each ptField In PT.PivotFields
        For Each vSubTotal In ptField.Subtotals
            vSubTotal = False
        Next
    Next
    TogglePTSubtotalsOff = True
errsub:
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Excel Auto Filters
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function ApplyAutoFilter(rngColumns As Excel.Range, vFilterFieldIndex As Variant, _
    vCritiria_1 As Variant, Optional xlOperator As XlAutoFilterOperator, Optional vCritiria_2 As Variant = -1) As Excel.Range
'Pass in selected range for filter (including headers), Passes back data range based on AutoFilter results (without headers).
'For tips on filter options: http://www.ozgrid.com/News/jul-2006.htm
'Returns "nothing" if no matching filtered range is found.
'Can return multiple non-contiguous ranges.
'Only tested for columns

    Dim rng As Excel.Range

    Application.DisplayAlerts = False
    rngColumns.Parent.AutoFilterMode = False

    'Apply AutoFilter
    If xlOperator = 0 And vCritiria_2 = -1 Then
        Call rngColumns.AutoFilter(vFilterFieldIndex, vCritiria_1)
    ElseIf xlOperator <> 0 And vCritiria_2 <> -1 Then
        Call rngColumns.AutoFilter(vFilterFieldIndex, vCritiria_1, xlOperator, vCritiria_2)
    Else
        'Call MsgBox("Input Error.  Funtion will now exit.", vbCritical)
        GoTo errsub
    End If

    On Error Resume Next

    Set ApplyAutoFilter = rngColumns.Parent.Range(rngColumns.Parent.Cells(rngColumns.Row + 1, rngColumns.Column), rngColumns.Parent.Cells(rngColumns.End(xlDown).Row, rngColumns.End(xlToRight).Column)).SpecialCells(xlCellTypeVisible)

    'Could also use RowHeight to determine which cells are "visible" in range after filter.
    'ActiveCell.RowHeight = 0

    '"Subscript out of range" error possible
    If Err.Number <> 0 Then
        Set ApplyAutoFilter = Nothing
        'err.clear
    End If

    'Exit Function
errsub:
    rngColumns.Parent.AutoFilterMode = False 'Turn filter off
    Application.DisplayAlerts = True
End Function

'Private Function Auto_Filter_Return_Column_Range(rngStartCell As Excel.Range) As Excel.Range
'Has been seen to return false positives, when UsedRange returns range larger than range that is filtered.
'    Dim rngTemp As Excel.Range
'    Dim rng As Excel.Range
'    Dim rngResult As Excel.Range
'    Dim lCount As Long
'
'    Set rngTemp = Intersect(rngStartCell.EntireColumn, rngStartCell.Parent.UsedRange)
'    Set rngTemp = rngStartCell.Parent.Range(rngStartCell.Offset(1, 0).Address, rngTemp(rngTemp.Cells.Count).Address) '(xlDown).Address)
'
'    On Error Resume Next
'    Set rngTemp = rngTemp.SpecialCells(xlCellTypeVisible) 'Will error if no cells are found.
'
'    If Err Then
'        Set rngTemp = Nothing
'    Else
'        Set Auto_Filter_Return_Column_Range = rngTemp
'    End If
'
'    On Error GoTo 0
'    Err.Clear
'End Function

'Only tested for columns
'Range("C5:C56").AdvancedFilter Action:=xlFilterInPlace, Unique:=True
'Parameters are shown as "variant" in help?  Do they need to be?
'Need to put case statement in here.
'Private Function ApplyAdvancedFilter(xlAction As XlFilterAction, Optional rngSource As Excel.Range, Optional rngDestination As Excel.Range, Optional bUnique As Boolean = False) As Excel.Range
''Function ApplyAdvancedFilter(rngColumns As Excel.Range, xlAction As XlFilterAction, rngSource As Excel.Range, Optional rngDestination As Excel.Range = rngSource, Optional bUnique As Boolean = False) As Excel.Range
'    Dim rng As Excel.Range
'
'    Application.DisplayAlerts = False
'
'    If rngSource Is Nothing Then
'        Set rngSource = Selection
'    End If
'    If rngDestination Is Nothing Then
'        Set rngDestination = Selection
'    End If
'
'    'Turn Filter Off - Need to keep track of what was showing before?
'    'rngSource.Parent.Name .ShowAllData
'    ActiveSheet.ShowAllData 'Turn off AdvanceFilters (will error if already off)
'    rngSource.Parent.ShowAllData 'Turn off AdvanceFilters (will error if already off)
'    'Worksheets("Percentage_of_Work").ShowAllData
'
'    'Apply AdvancedFilter
'    If xlAction = xlFilterInPlace Then 'bUnique rngDestination
'        Set rngDestination = rngSource
'    ElseIf xlAction = xlFilterCopy And rngDestination.Address = rngSource.Address Then
'    'ElseIf xlAction = xlFilterCopy And rngDestination Is rngSource Then
'        'Call MsgBox("A destination range must be supplied to perform copy.  Funtion will now exit.", vbCritical)
'        GoTo Errsub
'    End If
'
'    Call rngSource.AdvancedFilter(xlAction, rngSource, rngDestination, bUnique)
'    'Range("C5:C56").AdvancedFilter Action:=xlFilterInPlace, Unique:=True
'
'    On Error Resume Next
'
'    Set ApplyAdvancedFilter = Excel.Range(Cells(rngSource.Row + 1, rngSource.Column), Cells(rngSource.End(xlDown).Row, rngSource.End(xlToRight).Column)).SpecialCells(xlCellTypeVisible)
'
'    '"Subscript out of range" error possible
'    If Err.Number <> 0 Then
'        Set ApplyAdvancedFilter = Nothing
'    End If
'
'Errsub:
'    rngSource.Parent.ShowAllData 'Turn filter off
'    Application.DisplayAlerts = True
'End Function

Private Function UniqueItemsInRange(rngIn As Excel.Range, cellProp As cellProperty, bSkipBlanks As Boolean, Optional Count As Variant) As Excel.Range
'Passes back range of unique items based on CellProperty
    Dim rng As Excel.Range, rng2 As Excel.Range
    Dim rngResult As Excel.Range
    Dim bFound As Boolean

'   Loop through the input array
    If Not rngIn Is Nothing Then
'        If rngIn.Cells.Count > 1 Then
            For Each rng In rngIn   'Loop through range passed in.
                If Not rngResult Is Nothing Then
                    For Each rng2 In rngResult
                        Select Case cellProp
                            Case cellProperty.Value
                                If (rng2.Value = rng.Value) Or (bSkipBlanks = True And rng.Value = Empty) Then
                                    bFound = True
                                    Exit For
                                End If
                        End Select
                    Next rng2
                End If
                If bFound = False Then
                    Set rngResult = AddRange(rngResult, rng)
                Else
                    bFound = False
                End If
            Next rng
'            Next Element
'        End If
    End If

    If IsMissing(Count) = False Then
        Count = rngResult.Cells.Count
    End If
    Set UniqueItemsInRange = rngResult

'    'Could also return an array such as: (though this doesn't appear to work for nonconfiguous ranges) see VariantToArray()
'    If Not rngResult Is Nothing Then
'        Dim arry As Variant
'        arry = rngResult.Value
'    End If
End Function

Private Function UniqueItemsInRangeByValue(rngSource As Excel.Range, bSkipBlanks As Boolean) As Excel.Range
'Passes back range of unique items based on value
    Dim rng As Excel.Range
    Dim rngResult As Excel.Range

    If Not rngSource Is Nothing Then
        For Each rng In rngSource
            If (IsEmpty(rng) = True And bSkipBlanks = False) Or (IsEmpty(rng) = False) Then
                If rngResult Is Nothing Then 'First one
                    Set rngResult = rng
                ElseIf rngResult.Find(rng.Value, , , Excel.XlLookAt.xlWhole) Is Nothing Then 'Is already in destination set?
                    Set rngResult = Application.Union(rngResult, rng)
                End If
            End If
        Next
    End If
    Set UniqueItemsInRangeByValue = rngResult
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Chart Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Function BuildChart(vChartType As XlChartType, rngSource As Excel.Range, wsActiveWorkSheet As Worksheet, vPlotBy As XlRowCol, _
    strChartTitle As String, strCategoryAxisTitle As String, strValueAxisTitle As String, bHasLegend As Boolean, _
    vLegendPos As XlLegendPosition, vSeriesNames As Variant, vSeriesLineWeight As XlBorderWeight, rPos As RECT, Optional chrtToReplace As Chart) As Chart
    'Initially created for xlXYScatterLinesNoMarkers with series based on columns passed in which are in the format X | Y |X | Y...
    'Add more as required.

    Dim CT As ChartObject
    Dim i As Long
    Dim tSeries As Series
    Dim tmpRange As Excel.Range

    'Remove previous chart.
    If Not chrtToReplace Is Nothing Then
        Call RemovePreviousCharts(wsActiveWorkSheet, chrtToReplace)
    End If

    'Add new chart
    Set CT = ActiveSheet.ChartObjects.Add(rPos.Left, rPos.Top, rPos.Right - rPos.Left, rPos.Bottom - rPos.Top) 'Left, Top, Width, Height

    'Set CT = ActiveSheet.ChartObjects.Add(rPosition.Left, rPosition.Right - rPosition.Left, rPosition.Top, rPosition.Bottom - rPosition.Top)
    With CT.Chart
        .ChartType = vChartType
        .SetSourceData rngSource, vPlotBy 'Default series are created here.
        '.Location xlLocationAsObject, strSheetName 'Don't need and causes crash.
        .HasTitle = IIf(strChartTitle = "", False, True)
        .ChartTitle.Characters.Text = strChartTitle
        .Axes(xlCategory, xlPrimary).HasTitle = IIf(strCategoryAxisTitle <> vbNullString, True, False)
        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = strCategoryAxisTitle
        .Axes(xlValue, xlPrimary).HasTitle = IIf(strValueAxisTitle <> vbNullString, True, False)
        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = strValueAxisTitle
        .HasLegend = bHasLegend
        .Legend.Position = vLegendPos

        'Delete all previous series(created by default when chart is created above)
        Call RemoveChartSeries(CT.Chart)

        'Create series based on columns passed in which are in the format X | Y |X | Y...
        If rngSource.Cells.Count > 0 Then 'Check that there is a range
            Set tmpRange = rngSource

            For i = 1 To rngSource.Columns.Count Step 2 'Loop for each pair
                'If series box is check and there are values to plot in the series.
                If vSeriesNames(((i + 1) / 2) - 1, 1) And tmpRange.Cells(1).Value <> 0 Then
                    Set tmpRange = Excel.Range(tmpRange.Cells(1), tmpRange.End(xlDown))
                    With .SeriesCollection.NewSeries
                        'Narrow these down to used cells only.
                        .Values = tmpRange.Offset(0, 1).Value
'                        .Values = "=" & rngSource.Parent.Name & "!" & tmpRange.Offset(0, 1).Address(, , xlR1C1)
                        .XValues = tmpRange.Value
'                        .XValues = "=" & rngSource.Parent.Name & "!" & tmpRange.Address(, , xlR1C1)
                        .name = vSeriesNames(((i + 1) / 2) - 1, 0)
                        .Border.Weight = vSeriesLineWeight
                        '.Border.LineStyle = xlNone ' Hide series
                    End With
                End If
                Set tmpRange = tmpRange.Offset(0, 2)
            Next i

            'Add 1 default series otherwise chart disappears.
            If .SeriesCollection.Count = 0 Then
                Set tSeries = .SeriesCollection.NewSeries
                With tSeries
                    .name = "No Data Selected"
                    .Border.Weight = vSeriesLineWeight
                End With
            End If
        End If
    End With

    Set BuildChart = CT.Chart
'--------------------------------------
    'Add new chart
'    Charts.Add

'    ActiveChart.ChartType = vChartType
'    ActiveChart.SetSourceData rngSource, vPlotBy
'    ActiveChart.Location xlLocationAsObject, strSheetName 'Where the chart is inserted on the sheet.

'    With ActiveChart 'Not sure why, but "With" doesn't always work.
'        .HasTitle = True
'        .ChartTitle.Characters.Text = strChartTitle
'        .Axes(xlCategory, xlPrimary).HasTitle = IIf(strCategoryAxisTitle <> vbNullString, True, False)
'        .Axes(xlCategory, xlPrimary).AxisTitle.Characters.Text = strCategoryAxisTitle
'        .Axes(xlValue, xlPrimary).HasTitle = IIf(strValueAxisTitle <> vbNullString, True, False)
'        .Axes(xlValue, xlPrimary).AxisTitle.Characters.Text = strValueAxisTitle
'        .HasLegend = bHasLegend
'        .Legend.Position = vLegendPos
'    End With

    'Delete all previous series(created by default when chart is created above)
'    Call RemoveUnwantedChartSeries(ActiveChart)

'    'Create series based on columns passed in which are in the format X | Y |X | Y...
'    If rngSource.Cells.Count > 0 Then 'Check that there is a range
'        Set tmpRange = rngSource
'
'        For i = 1 To rngSource.Columns.Count Step 2 'Loop for each pair
'            'If series box is check and there are values to plot in the series.
'            If vSeriesNames(((i + 1) / 2) - 1, 1) And tmpRange.Cells(1).Value <> 0 Then
'                Set tmpRange = Excel.Range(tmpRange.Cells(1), tmpRange.End(xlDown))
'                Set tSeries = ActiveChart.SeriesCollection.NewSeries
'                With tSeries
'                    .Name = vSeriesNames(((i + 1) / 2) - 1, 0)
'                    'Narrow these down to used cells only.
'                    .XValues = "=" & rngSource.Parent.Name & "!" & tmpRange.Address(, , xlR1C1)
'                    .Values = "=" & rngSource.Parent.Name & "!" & tmpRange.Offset(0, 1).Address(, , xlR1C1)
'                    .Border.Weight = vSeriesLineWeight
'                    '.Border.LineStyle = xlNone ' Hide series
'                End With
'            End If
'            Set tmpRange = tmpRange.Offset(0, 2)
'        Next i
'
'        'Add default series otherwise chart disappears.
'        If ActiveChart.SeriesCollection.Count = 0 Then
'            Set tSeries = ActiveChart.SeriesCollection.NewSeries
'            With tSeries
'                .Name = "No Data Selected"
'                .Border.Weight = vSeriesLineWeight
'            End With
'        End If
'    End If
'    Set BuildChart = ActiveChart
End Function

Private Sub RemovePreviousCharts(tWS As Worksheet, Optional tRemoveSingleChart As Chart)
'Remove all charts on a worksheet, or just one if passed in.
    Dim CT As ChartObject
    If tRemoveSingleChart Is Nothing Then
        tWS.ChartObjects.Delete
'        For Each CT In tWS.ChartObjects
'        'For Each CT In ActiveSheet.ChartObjects
'            CT.Delete
'        Next CT
    Else
        For Each CT In tWS.ChartObjects
            If CT.Chart Is tRemoveSingleChart Then
                CT.Delete
                Exit For
            End If
        Next CT
    End If
End Sub

Private Sub EmbeddedChartFromScratch()
'Untested
    Dim myChtObj As ChartObject
    Dim rngChtData As Excel.Range
    Dim rngChtXVal As Excel.Range
    Dim iColumn As Long

    ' make sure a range is selected
    If TypeName(Selection) <> "Range" Then Exit Sub

    ' define chart data
    Set rngChtData = Selection

    ' define chart's X values
    With rngChtData
        Set rngChtXVal = .Columns(1).Offset(1).Resize(.Rows.Count - 1)
    End With

    ' add the chart
    Set myChtObj = ActiveSheet.ChartObjects.Add _
        (Left:=250, Width:=375, Top:=75, Height:=225)
    With myChtObj.Chart

        ' make an XY chart
        .ChartType = xlXYScatterLines

        ' remove extra series
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete
        Loop

        ' add series from selected range, column by column
        For iColumn = 2 To rngChtData.Columns.Count
            With .SeriesCollection.NewSeries
                .Values = rngChtXVal.Offset(, iColumn - 1)
                .XValues = rngChtXVal
                .name = rngChtData(1, iColumn)
            End With
        Next

    End With
End Sub

Private Sub ResizeAndRepositionChart(rChart As Chart, rPosition As RECT)
'Untested
' The ChartObject is the Chart's parent
    With rChart.Parent
        .Left = rPosition.Left
        .Width = rPosition.Right - rPosition.Left
        .Top = rPosition.Top
        .Height = rPosition.Bottom - rPosition.Top
    End With
End Sub

Private Function BuildChartSeriesArrays(TargetChart As Excel.Chart, FieldHeader As Excel.Range, DateRange As Excel.Range)
    'Example of how to create dynamic array to use with AddChartSeriesByArray().
    Dim i As Long
    Dim rngDate As Excel.Range
    Dim rngField As Excel.Range
    Dim YValues() As Double 'Percent Change
    Dim XValues() As Double 'Dates
    Dim dTemp As Double
    Dim dBenchmark As Double

'    If Not FieldHeader Is Nothing And Not DateRange Is Nothing Then 'Nothing selected
'        For Each rngField In FieldHeader.Cells 'Columns
'            ReDim YValues(DateRange.Cells.Count - 1)
'            ReDim XValues(DateRange.Cells.Count - 1)
'            For Each rngDate In DateRange.Cells 'Rows
'                XValues(i) = CDbl(rngDate.Value) 'Dates
'                dTemp = CDbl(Application.Intersect(rngDate.EntireRow, rngField.EntireColumn).Value)
'                If i = 0 Then 'Calculate Percent Change
'                    YValues(i) = 0
'                    dBenchmark = dTemp
'                Else
'                    YValues(i) = dTemp / dBenchmark - 1 'Percent Change
'                End If
'                i = i + 1
'            Next
'            Call AddChartSeriesByArray(TargetChart, rngField.Value, XValues, YValues)
'            i = 0
'        Next
'    End If
End Function

Private Function AddChartSeriesByArray(TargetChart As Excel.Chart, SeriesName As String, Rows As Variant, Fields As Variant)
'Add chart series that points to array rather than range source.
'Passed in Variants are single dimension.
    With TargetChart.SeriesCollection.NewSeries
        .Values = Fields 'Y Values
        .XValues = Rows 'X values
        .name = SeriesName 'Value Header
        '.Border.Weight = vSeriesLineWeight
        '.Border.LineStyle = xlNone ' Hide series
    End With
End Function

Private Sub RemoveChartSeries(rChart As Chart)
    With rChart
'        rChart.ChartArea.Select
        Do Until .SeriesCollection.Count = 0
            .SeriesCollection(1).Delete 'Delete first item in collection.
        Loop
    End With

'    Dim tSeries As Series
'    For Each tSeries In ActiveChart.SeriesCollection
'        tSeries.Delete
'    Next tSeries
End Sub

Public Function InsertTextToTextBox(objTextBox As Excel.Shape, strCharacters As String)
'Insert Characters Into TextBox created with .Shapes.AddTextbox
'Workaround for known issue wherein on 255 characters can be added at one time when setting objTextBox.TextFrame.Characters.Text() in Excel 2003.
'Formatting of particular text in textbox is performed such as:
'objTextBox.TextFrame.Characters(1, Len("Projects:")).Font.Bold = True
'objTextBox.TextFrame.Characters(InStr(1, m_StackedAreaInfo, "Resources:"), Len("Resources:")).Font.Bold = True

    Dim Index As Long
    Dim Value As String
'    Dim bPreviousScreenUpdate As Boolean

'    bPreviousScreenUpdate = Application.ScreenUpdating
'    Application.ScreenUpdating = False
    With objTextBox.TextFrame
        If Application.Version >= 14 Then '2010
            .Characters.Text = vbNullString
            .Characters.Text = strCharacters
        Else
            If .Characters.Count > 0 Then
                .Characters(1, .Characters.Count).Text = vbNullString
            End If

            For Index = 0 To Int(Len(strCharacters) / 255)
                If .Characters.Count = 0 Then
                    .Characters.Text = VBA.Mid(strCharacters, (Index * 255) + 1, 255)
                Else
                    Call .Characters(.Characters.Count + 1).Insert(VBA.Mid(strCharacters, (Index * 255) + 1, 255))
                End If
            Next Index
        End If

        If .Characters.Count > 0 Then
            .Characters(1, .Characters.Count).Font.Bold = False
        End If
    End With
'   Application.ScreenUpdating = bPreviousScreenUpdate
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Range functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetAverageOfUsedCellsInColumn(rngInput As Excel.Range) As Double
'Used in function tables.
    Dim rngNew As Excel.Range
    Dim rngResult As Excel.Range

    If IsEmpty(rngInput.Offset(1, 0)) = False Then 'Simple check for empty value.
        Set rngNew = Excel.Range(rngInput.Offset(1, 0), GetLastUsedCellInColumnByStartCell(rngInput))
        GetAverageOfUsedCellsInColumn = Application.WorksheetFunction.Average(rngNew)
    End If
End Function

Private Function AddRange(ByVal rngBase As Excel.Range, ByVal rngAdd As Excel.Range) As Excel.Range ', ByVal bUniqueCells As Boolean) As Excel.Range
'Use For Each Range in Range.cells to extract values so that
'Doesn't care if cells are merged or if they use ":" notation.

'The following function will return a Range object that is the logical union of two ranges.  Unlike the Application.Union method, AddRange will not return duplicate cells in the result.  For example,
'Application.Union(Range("A1:B3"), Range("B3:D5")).Cells.Count
'will return 15, since B3 is counted twice, once in each range.
'AddRange(Range("A1:B3"), Range("B3:D5")).Cells.Count will return 14, counting B3 only once.
'- Modified from original to include Unique cells switch
    Dim rng As Excel.Range

    If rngBase Is Nothing Then
        If rngAdd Is Nothing Then
            Set AddRange = Nothing
        Else
            Set AddRange = rngAdd
        End If
    Else
        If rngAdd Is Nothing Then
            Set AddRange = rngBase
        Else
            Set AddRange = rngBase
            For Each rng In rngAdd
                If Application.Intersect(rng, rngBase) Is Nothing Then 'If not already in range.
'                    'Returns individual cells. No ":" notation usage.  Usually not required as for each Range in Range.cells used to extract values.
'                    If bUniqueCells = True Then
'                        Set AddRange = Range(AddRange.Address & "," & rng.Address)
'                    Else
                        Set AddRange = Application.Union(AddRange, rng) 'Returns merged ranges of cells using ":" notation.
'                    End If
                End If
            Next rng
        End If
    End If
End Function

Private Function RemoveRange(ByVal rngBase As Excel.Range, ByVal rngRemove As Excel.Range) As Excel.Range
'http://www.dailydoseofexcel.com/archives/2007/08/17/two-new-range-functions-union-and-subtract/
'Remove a single range value from a base range and passes back new range.
    Dim rngNew As Excel.Range
    Dim rng As Excel.Range

    On Error Resume Next
    Set rngRemove = rngRemove(1) 'Make sure only a single range
    'Not included in base range
    If Application.Intersect(rngBase, rngRemove).Address <> rngRemove.Address Then
        Set rngNew = rngBase
    Else
        On Error GoTo 0
        'Approach is to rebuild a range skipping items in rngRemove.
        For Each rng In rngBase
            If Application.Intersect(rng, rngRemove) Is Nothing Then 'If not in rngRemove, (don't want to remove it).
                Set rngNew = AddRange(rngNew, rng) 'ReAdd each item that are not in rngRemove.
            End If
        Next rng
    End If
    Set RemoveRange = rngNew
End Function

Private Function RemoveEmptyRowsAndColumnsFromRange(rngSource As Excel.Range) As Excel.Range
    Set rngSource = RemoveEmptyRowsFromRange(rngSource)
    Set rngSource = RemoveEmptyColumnsFromRange(rngSource)
    Set RemoveEmptyRowsAndColumnsFromRange = rngSource
End Function

Private Function RemoveEmptyRowsFromRange(rngSource As Excel.Range) As Excel.Range
'Remove empty rows from range object.
'Does not manipulate the actual range on the worksheet.
'CountBlank does not seem to function with non-contiguous ranges.  Possibly add "Areas" aspect.
'Slow!
    Dim rng As Excel.Range
    Dim rngColumn As Excel.Range
    Dim rngRemove As Excel.Range

    Set rngColumn = Application.Intersect(rngSource(1).EntireColumn, rngSource)

    For Each rng In rngColumn
        If WorksheetFunction.CountBlank(Application.Intersect(rng.EntireRow, rngSource)) = rngSource.Columns.Count Then
            Set rngRemove = AddRange(rngRemove, Application.Intersect(rng.EntireRow, rngSource))
        End If
    Next

    Set RemoveEmptyRowsFromRange = RemoveRange(rngSource, rngRemove)
End Function

Private Function RemoveEmptyColumnsFromRange(rngSource As Excel.Range) As Excel.Range
'Remove empty columns from range object.
'Does not manipulate the actual range on the worksheet.
'CountBlank does not seem to function with non-contiguous ranges.  Possibly add "Areas" aspect.
'Slow!
    Dim rng As Excel.Range
    Dim rngRow As Excel.Range
    Dim rngRemove As Excel.Range

    Set rngRow = Application.Intersect(rngSource(1).EntireRow, rngSource)

    For Each rng In rngRow
        If WorksheetFunction.CountBlank(Application.Intersect(rng(1).EntireColumn, rngSource)) = rngSource.Rows.Count Then
            Set rngRemove = AddRange(rngRemove, Application.Intersect(rng.EntireColumn, rngSource))
        End If
    Next

    Set RemoveEmptyColumnsFromRange = RemoveRange(rngSource, rngRemove)
End Function

Private Function GetRangeColumnLetter(rng As Excel.Range) As String
'Largest column is currently IV as of 2007
    Dim strResult As String

    Set rng = rng.Cells(1) 'Ensure that only one column is being considered in the event that more are passed in.
    strResult = rng.Address(True, True, xlA1)
    strResult = VBA.Strings.Mid$(strResult, 2, InStrRev(strResult, "$") - 2)
    GetRangeColumnLetter = strResult
    'IIf(rng.Column \ 64 > 0, VBA.Chr((rng.Column \ 64) + 64), "") & IIf(rng.Column Mod 64 > 0, VBA.Chr((rng.Column Mod 64) + 64), "")
End Function

Private Function RangeToVariant(rngSource As Excel.Range) As Variant
'Converts range into variant array

'Removed: doesn't support non-contiguous ranges.  Original Walkenbach function.
'    Dim arryNew As Variant
'    arryNew = Range(rngSource.Address).Value
'    Set RangeToVariant = arryNew

    Dim lCount As Long
    Dim rng As Excel.Range
    Dim aryUnique() As Variant
    ReDim Preserve aryUnique(rngSource.Count - 1) '1- rngSource.count

    For Each rng In rngSource
        aryUnique(lCount) = rng.Value
        lCount = lCount + 1
    Next

    RangeToVariant = aryUnique
End Function

Private Function RangeArrayToArray(vArray As Variant) As Variant
'Converts range array (created with rng.value) in one dimensional, zero based array.
'Currently only supports columns
'Can also call TransposeArray() as needed.
    Dim X As Long
    Dim Xupper As Long
    Dim XLower As Long
    Dim TempArray As Variant
    Dim lOffset As Long

    If NumberOfArrayDimensions(vArray) = 2 Then
        If UBound(vArray, 1) > UBound(vArray, 2) Then 'Columns
            lOffset = LBound(vArray, 1)
            Xupper = UBound(vArray, 1)
            XLower = LBound(vArray, 1)

            ReDim TempArray(0 To Xupper - lOffset)
            For X = XLower To Xupper
                TempArray(X - lOffset) = vArray(X, 1)
            Next
        Else 'Rows
            lOffset = LBound(vArray, 2)
            Xupper = UBound(vArray, 2)
            XLower = LBound(vArray, 2)

            ReDim TempArray(0 To Xupper - lOffset)
            For X = XLower To Xupper
                TempArray(X - lOffset) = vArray(1, X)
            Next
        End If
    End If

    RangeArrayToArray = TempArray
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

Private Function CellCountByColor(InRange As Excel.Range, lColor As Long, Optional OfText As Boolean = False) As Long
' This function return the number of cells in InRange with
' a background color, or if OfText is True a font color,
' equal to WhatColor.
'http://www.cpearson.com/excel/colors.htm
'Changed ColorIndex to Long representing RGB() value rather than index color
    Dim rng As Excel.Range
'    Application.Volatile True

    For Each rng In InRange.Cells
        If OfText = True Then
            CellCountByColor = CellCountByColor - (rng.Font.Color = lColor)
            'CellCountByColor = CellCountByColor - (rng.Font.ColorIndex = WhatColorIndex)
        Else
            CellCountByColor = CellCountByColor - (rng.Interior.Color = lColor)
            'CellCountByColor = CellCountByColor - (rng.Interior.ColorIndex = WhatColorIndex)
        End If
    Next rng
End Function

Private Function RangeOfColor(InRange As Excel.Range, lColor As Long, bIsText As Boolean) As Excel.Range
' This function returns a Range of cells in InRange with a background color, '
' or if OfText is True a font color, equal to WhatColor.
'Modified so that lColor represents RGB() color rather than index color.
    Dim rng As Excel.Range
'    Application.Volatile True
    For Each rng In InRange.Cells
        If bIsText = True Then
            If (rng.Font.Color = lColor) = True Then
                Set RangeOfColor = AddRange(RangeOfColor, rng)
            End If
        Else
            If (rng.Interior.Color = lColor) = True Then
                Set RangeOfColor = AddRange(RangeOfColor, rng)
            End If
        End If
    Next rng
End Function

Private Function CellColorIndex(InRange As Excel.Range, Optional OfText As Boolean = False) As Long
' This function returns the ColorIndex value of a the Interior
' (background) of a cell, or, if OfText is true, of the Font in the cell.
'The following function will return the ColorIndex property of a cell.
'Application.Volatile True
    Set InRange = Excel.Range("A6")
    If OfText = True Then
        CellColorIndex = InRange(1, 1).Font.ColorIndex
    Else
        CellColorIndex = InRange(1, 1).Interior.ColorIndex
    End If

'Possible XLColorIndex return values:
'White = 2
'Black = 1
'xlColorIndexNone = -4142 'or xlNone
'xlColorIndexAutomatic = -4105
End Function

Private Sub DeleteRangeExtents(rngAnchor As Excel.Range)
'Deletes range beginning with anchor and ending at end of sheet.  Used to remove non-visible items on sheet that are recognized by UsedRange.
'Implemented originally to clean ProModel outputs.
    On Error Resume Next 'Will error if no values in column.
        rngAnchor.Parent.Range(GetLastUsedCellInColumnByStartCell(rngAnchor).Offset(1, 0), GetLastCellInWorksheet(rngAnchor.Parent)).Delete (xlUp)
        rngAnchor.Parent.Range(GetLastUsedCellInRowByStartCell(rngAnchor).Offset(0, 1), GetLastCellInWorksheet(rngAnchor.Parent)).Delete (xlToLeft)
    On Error GoTo 0
End Sub

Private Function DeleteUnusedRowsInArea(rngSourceArea As Excel.Range, Optional bEntireRow As Boolean = True) As Excel.Range
    'Used to remove blank cell rows that cause file sizes to become larger.
    'Search backwards by rows to find first value in rngSource area, offset one row down, then delete included rngSource area 'under' found value.
    '
    'Delete entire rows is faster.
    Dim wksParent As Excel.Worksheet
    Dim rng As Excel.Range
    Dim strTemp As String
    
    Set wksParent = rngSourceArea.Parent
    Set rng = GetLastUsedCellInRangeAreaByStartCell(rngSourceArea)
    If Not rng Is Nothing Then
        Set rng = wksParent.Range(wksParent.Cells(rng.Offset(1).Row, rngSourceArea.Columns(1).Column), rngSourceArea.Cells(rngSourceArea.Rows.Count, rngSourceArea.Columns.Count))
        
        If bEntireRow = True Then
            Set rng = rng.EntireRow
        End If
        
        strTemp = rng.Address 'Store address that will shift, or be removed if included in delete.
        
'        Set rng = Application.Intersect(wksParent.UsedRange, rng) 'Is selection to delete in UsedRange.
        
        If Not rng Is Nothing Then
            Call rng.Delete(xlUp)
            Set DeleteUnusedRowsInArea = wksParent.Range(strTemp) 'Restore address
        End If
    End If
End Function

Private Function GetLastCellInRow(rngRow As Excel.Range) As Excel.Range
    'Last cell in row on worksheet
    Dim wksParent As Excel.Worksheet

    Set wksParent = rngRow.Parent
    Set rngRow = rngRow(1)
    Set GetLastCellInRow = Application.Intersect(rngRow.EntireRow, wksParent.Cells(wksParent.Rows.Count, wksParent.Columns.Count).EntireColumn)
End Function

Private Function GetLastCellInColumn(rngColumn As Excel.Range) As Excel.Range
    'Last cell in column on worksheet
    Dim wksParent As Excel.Worksheet

    Set wksParent = rngColumn.Parent
    Set rngColumn = rngColumn(1)
    Set GetLastCellInColumn = Application.Intersect(rngColumn.EntireColumn, wksParent.Cells(wksParent.Rows.Count, wksParent.Columns.Count).EntireRow)
End Function

Private Function GetLastCellInWorksheet(wksSource As Excel.Worksheet) As Excel.Range
'Likely $IV$65536 for 2003 xls
    Set GetLastCellInWorksheet = wksSource.Cells(wksSource.Rows.Count, wksSource.Columns.Count)
End Function

Private Function GetLastUsedCellInWorksheet(wksSource As Excel.Worksheet) As Excel.Range
    'Only considers used range.
    'UsedRange.Cells.Count are not always the last cell (right/bottom) in range
    Set GetLastUsedCellInWorksheet = wksSource.UsedRange.SpecialCells(xlCellTypeLastCell)
End Function

Private Function GetLastUsedCellInRowByStartCell(rngInput As Excel.Range) As Excel.Range
    'If row is empty or empty at and after start cell, passes back nothing
    Dim rng As Excel.Range

    Set rngInput = rngInput(1)
    Set rng = GetLastUsedCellInRowByValue(rngInput)
    If Not rng Is Nothing Then
        If rng.Column >= rngInput.Column Then
            Set GetLastUsedCellInRowByStartCell = rng
        End If
    End If
End Function

Private Function GetLastUsedCellInColumnByStartCell(rngInput As Excel.Range) As Excel.Range
    'If column is empty or empty at and after start cell, passes back nothing
    Dim rng As Excel.Range
    
    Set rngInput = rngInput(1)
    Set rng = GetLastUsedCellInColumnByValue(rngInput)
    If Not rng Is Nothing Then
        If rng.Row >= rngInput.Row Then
            Set GetLastUsedCellInColumnByStartCell = rng
        End If
    End If
End Function

Private Function GetLastUsedCellInRowByValue(rngStart As Excel.Range, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by rows.
    Dim rng As Excel.Range
    Set rng = rngStart.EntireRow.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious)
    If Not rng Is Nothing Then
'        If Not Application.Intersect(rngStart.EntireRow, rng) Is Nothing Then 'Will return different row if searching a blank area.
            Set GetLastUsedCellInRowByValue = rng
'        End If
    End If
End Function

Private Function GetLastUsedCellInColumnByValue(rngColumn As Excel.Range, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by columns.
    Dim rng As Excel.Range
    Set rng = rngColumn.EntireColumn.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious)
    If Not rng Is Nothing Then
'        If Not Application.Intersect(rngStart, rng) Is Nothing Then 'Will return different column if searching a blank area.
            Set GetLastUsedCellInColumnByValue = rng
'        End If
    End If
End Function

Private Function GetLastUsedCellInRangeAreaByStartCell(rngInput As Excel.Range) As Excel.Range
    'If empty or after start cell, passes back nothing
    Dim rng As Excel.Range

    Set rng = GetLastUsedCellInColumnByValue(rngInput)
    If Not rng Is Nothing Then
        If rng.Row >= rngInput.Row Then
            Set GetLastUsedCellInRangeAreaByStartCell = rng
        End If
    End If
End Function

Function GetLastUsedCell(wksSource As Excel.Worksheet) As Excel.Range
    'Error-handling here in case there is not any data in the worksheet
    On Error GoTo errsub:
    Set GetLastUsedCell = wksSource.Cells(GetLastUsedRow(wksSource).Row, GetLastUsedColumn(wksSource).Column)
errsub:
End Function

Private Function GetLastUsedRow(wksSource As Excel.Worksheet, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by Rows.
    On Error GoTo errsub:
    Set GetLastUsedRow = wksSource.UsedRange.Cells.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious)
errsub:
End Function

Private Function GetLastUsedColumn(wksSource As Excel.Worksheet, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by Columns.
    On Error GoTo errsub:
    Set GetLastUsedColumn = wksSource.UsedRange.Cells.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlPrevious)
errsub:
End Function

Private Function GetFirstUsedCellInColumnByStartCell(rngInput As Excel.Range) As Excel.Range
    'If column is empty or empty at and after start cell, passes back nothing
    Dim rng As Excel.Range

    Set rngInput = rngInput(1)
    Set rng = GetFirstUsedCellInRowByValue(rngInput)
    If Not rng Is Nothing Then
        Set GetFirstUsedCellInColumnByStartCell = rng
    End If
End Function

Private Function GetFirstUsedCellInRowByStartCell(rngInput As Excel.Range) As Excel.Range
    'If row is empty or empty at and after start cell, passes back nothing
    Dim rng As Excel.Range

    Set rngInput = rngInput(1)
    Set rng = GetFirstUsedCellInRowByValue(rngInput)
    If Not rng Is Nothing Then
        Set GetFirstUsedCellInRowByStartCell = rng
    End If
End Function

Private Function GetFirstUsedCellInRowByValue(rngStart As Excel.Range, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by rows.
    Dim rng As Excel.Range
    Set rng = rngStart.EntireRow.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext)
    If rng.Row = rngStart.Row Then 'Will return different column if searching a blank area.
        Set GetFirstUsedCellInRowByValue = rng
    End If
End Function

Private Function GetFirstUsedCellInColumnByValue(rngStart As Excel.Range, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching forwards by columns.
    Dim rng As Excel.Range
    Set rng = rngStart.EntireColumn.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
    If rng.Column = rngStart.Column Then 'Will return different column if searching a blank area.
        Set GetFirstUsedCellInColumnByValue = rng
    End If
End Function

Private Function GetFirstUsedRow(wksSource As Excel.Worksheet, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by Rows.
    On Error GoTo errsub:
    Set GetFirstUsedRow = wksSource.UsedRange.Cells.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext)
errsub:
End Function

Private Function GetFirstUsedColumn(wksSource As Excel.Worksheet, Optional vValue As Variant = "*") As Excel.Range
    'Search for any entry, by searching backwards by Columns.
    On Error GoTo errsub:
    Set GetFirstUsedColumn = wksSource.UsedRange.Cells.Find(vValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext)
errsub:
End Function

Private Function GetUsedRangeByStartCell(rngAnchorCell As Excel.Range) As Excel.Range
'Return only the range of visible items within the UsedRange given the anchor point provided.
'Assumes that range is the same height as starts rngAnchorrange.
'UsedRange includes cells that have equations or other formatting that is not necessarily visible.
    Set GetUsedRangeByStartCell = rngAnchorCell.Parent.Range(rngAnchorCell(1), rngAnchorCell.Parent.Range(GetLastUsedCellInRowByStartCell(rngAnchorCell(1)), GetLastUsedCellInColumnByStartCell(rngAnchorCell(1))))
End Function

Private Function GetUsedArea(rngSource As Excel.Range) As Excel.Range
'Return used range within passed in range.
    Dim rng As Excel.Range
    Dim lRow As Long
    Dim lCol As Long

    lRow = rngSource(1).Row
    lCol = rngSource(1).Column

'    For Each rng In rngSource
    Set rng = GetUsedRangeByStartCell(rngSource)
'    Set rng = GetLastUsedCellInColumnByStartCell(rng)
    If Not rng Is Nothing Then
        If lRow < rng.Row Then
            lRow = rng.Row
        End If
        If lCol < rng.Column Then
            lCol = rng.Column
        End If
    End If
'    Next

    If lRow > 0 And lCol > 0 Then
        Set GetUsedArea = rngSource.Parent.Range(rngSource(1), rngSource.Parent.Cells(lRow, lCol))
    End If
End Function

Private Function GetUsedRow(ByVal rngInput As Excel.Range) As Excel.Range
'Returns range of used row.
'Returns nothing if no range.
    On Error GoTo errsub
    Dim rngStart As Excel.Range
    Dim rngEnd As Excel.Range

    Set rngInput = rngInput(1)
    Set rngStart = GetFirstUsedCellInRowByStartCell(rngInput)
    Set rngEnd = GetLastUsedCellInRowByStartCell(rngInput)

    If Not rngStart Is Nothing And Not rngEnd Is Nothing Then
        Set GetUsedRow = rngInput.Parent.Range(rngStart, rngEnd)
    End If
errsub:
End Function

Private Function GetUsedColumn(ByVal rngInput As Excel.Range) As Excel.Range
'Returns range of used column.
'Returns nothing if no range.
    On Error GoTo errsub
    Dim rngStart As Excel.Range
    Dim rngEnd As Excel.Range

    Set rngInput = rngInput(1)
    Set rngStart = GetFirstUsedCellInColumnByStartCell(rngInput)
    Set rngEnd = GetLastUsedCellInColumnByStartCell(rngInput)

    If Not rngStart Is Nothing And Not rngEnd Is Nothing Then
        Set GetUsedColumn = rngInput.Parent.Range(rngStart, rngEnd)
    End If
errsub:
End Function

Private Function GetUsedColumnByStartCell(ByVal rngInput As Excel.Range) As Excel.Range
'Returns range of used column, starting with first cell in rngInput.
'Returns nothing if no range.
    On Error GoTo errsub
    Dim rng As Excel.Range

    Set rngInput = rngInput(1)
    Set rng = GetLastUsedCellInColumnByStartCell(rngInput) 'Checks if found cell is before rngInput
    If Not rng Is Nothing Then
        Set GetUsedColumnByStartCell = rngInput.Parent.Range(rngInput, rng)
    End If

errsub:
If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Private Function GetUsedRowByStartCell(ByVal rngInput As Excel.Range) As Excel.Range
'Returns range of used row, starting with first cell in rngInput.
'Returns nothing if no range.
    On Error GoTo errsub
    Dim rng As Excel.Range

    Set rngInput = rngInput(1)
    Set rng = GetLastUsedCellInRowByStartCell(rngInput) 'Checks if found cell is before rngInput
    If Not rng Is Nothing Then
        Set GetUsedRowByStartCell = rngInput.Parent.Range(rngInput, rng)
    End If

errsub:
If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Private Sub GetTwoColumnSourceData(ByVal rngCellSource As Excel.Range, ByVal rngCellDestination As Excel.Range)
    'Originally implemented for time Series Data where time/value pair is expected and time is assumed to not be 0.
    'Copies rngCellSource column into rngCellDestination for a pair of columns and then trims off 0's from the end of the data.

    Dim rngSource As Excel.Range
    Dim rngDestination As Excel.Range
    Dim rngTemp As Excel.Range
    Dim shtCalling As Worksheet

    Set shtCalling = rngCellDestination.Parent

    'Clear range of previous names and values.
    Excel.Range(rngCellDestination, rngCellDestination.Offset(shtCalling.UsedRange.Rows.Count + rngCellDestination.Row, 0)).Clear '2 extra cells cleared.
    'Time - Copy range
    Call CopyColumnRangeDynamic(rngCellSource, rngCellDestination, xlPasteValues)

    Set rngCellSource = rngCellSource.Offset(0, 1)
    Set rngCellDestination = rngCellDestination.Offset(0, 1)

    'Clear range of previous names and values.
    Excel.Range(rngCellDestination, rngCellDestination.Offset(shtCalling.UsedRange.Rows.Count + rngCellDestination.Row, 0)).Clear '2 extra cells cleared.
    'Count - Copy range
    Call CopyColumnRangeDynamic(rngCellSource, rngCellDestination, xlPasteValues)

    'Remove training zeros "0" from Time/Count dataset according to values in time column.
    Set rngCellDestination = rngCellDestination.Offset(0, -1)

    Set rngTemp = Excel.Range(rngCellDestination, rngCellDestination.Offset(shtCalling.UsedRange.Rows.Count + rngCellDestination.Row, 0))
    'Range(rngCellDestination, rngCellDestination.Offset(shtCalling.UsedRange.Rows.Count + rngCellDestination.Row, 0)).Select

    Set rngTemp = rngTemp.Find("0", rngCellDestination, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, False)
    '.Find may not work as expected if column is not wide enough to display results of a function in a cell.
    'Selection.Find("0", rngCellDestination, Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, False).Select
    'Warning - Will produce unexpected results if zeros are in the middle of the dataset.
    'If all zeros, one zero will be left in column.
    'This should not happen for valid time-series data as items can't change state at zero time.
    Excel.Range(rngTemp, rngTemp.Offset(0, 1).End(xlDown)).Clear
    'Range(ActiveCell, ActiveCell.Offset(0, 1).End(xlDown)).Clear

End Sub

Private Sub CopyColumnRangeDynamic(rngSourceStartCell As Excel.Range, rngDesinationStartCell As Excel.Range, xlPasteValueType As XlPasteType)
'Copies column of data based on start cell and used range to destination colummn based on start cell.
'Use Paste or param of Copy instead so that range doesn't need to be specified.
    Dim copyrange As Excel.Range
    Dim targetrange As Excel.Range
    Dim iRangeCount As Long

    iRangeCount = rngSourceStartCell.Parent.UsedRange.Rows.Count + rngSourceStartCell.Row   'Note:Includes a little extra range
    'iRangeCount = Worksheets(rngSourceStartCell.Parent.Name).UsedRange.Rows.Count + rngSourceStartCell.Row   'Note:Includes a little extra range

    Set copyrange = rngSourceStartCell.Parent.Range(rngSourceStartCell, rngSourceStartCell.Offset(iRangeCount))
    'Set copyrange = Worksheets(rngSourceStartCell.Parent.Name).Range(rngSourceStartCell, rngSourceStartCell.Offset(iRangeCount))

    Set targetrange = rngDesinationStartCell.Parent.Range(rngDesinationStartCell, rngDesinationStartCell.Offset(iRangeCount))
    'Set targetrange = Worksheets(rngDesinationStartCell.Parent.Name).Range(rngDesinationStartCell, rngDesinationStartCell.Offset(iRangeCount))

    Call copyrange.Copy
    Call targetrange.PasteSpecial(xlPasteValueType)

    'targetrange.Value = copyrange.Value
End Sub

Private Function FindCellRangesByColor(ByVal rngStart As Excel.Range, lColor As Long, Orientation As XlRowCol, bUniqueCells As Boolean) As Excel.Range
'Function FindCellRangesByColorIndex(ByVal rngStart As Excel.Range, lColorIndex As Long, Orientation As XlRowCol, bUniqueCells As Boolean) As Excel.Range
'Returns non-contiguous range of cells in a row/column of a given color.
'Selection.Interior.ColorIndex = 1 (Black)
    Dim rngLastCell As Excel.Range
    Dim rngUsed As Excel.Range
    Dim rngResult As Excel.Range
    Dim vRanges As Variant

    If Orientation = xlRows Then
        'Create range consisting of start = rngStart and end = last cell in usedrange in given row.
        Set rngUsed = Intersect(rngStart.Parent.UsedRange, rngStart.Rows(1).EntireRow) '.End(xlToRight)
        Set rngUsed = rngStart.Parent.Range(rngStart, rngStart.Parent.Cells(rngUsed.Row, rngUsed.Columns.Count))
    Else
        'Create range consisting of start = rngStart and end = last cell in usedrange in given column.
        Set rngUsed = Intersect(rngStart.Parent.UsedRange, rngStart.Columns(1).EntireColumn) '.End(xlDown)
        Set rngUsed = rngStart.Parent.Range(rngStart, rngStart.Parent.Cells(rngUsed.Rows.Count, rngUsed.Column))
    End If

    If CellCountByColor(rngUsed, lColor) > 0 Then 'Look for at least 1 bar in range.
        Set FindCellRangesByColor = RangeOfColor(rngUsed, lColor, bUniqueCells)
    Else
        FindCellRangesByColor = Nothing
    End If

End Function

Private Function InRange(rng1, rng2) As Boolean
'Returns True if rng1 is a subset of rng2
    If rng1.Parent.Parent.name = rng2.Parent.Parent.name Then
        If rng1.Parent.name = rng2.Parent.name Then
            If Union(rng1, rng2).Address = rng2.Address Then
                InRange = True
            End If
        End If
    End If
End Function

Private Function GetRangeIndexInRangeArray(rngMember As Excel.Range, rngParent As Excel.Range) As Long
'Returns 1 based index into a Range array that includes the supplied Member range.
    Dim lCount As Long
    Dim rngTest As Excel.Range

    If Not Intersect(rngMember, rngParent) Is Nothing Then 'And rngParent.Rows.Count = 1 Then
        For Each rngTest In rngParent
            lCount = lCount + 1
            If rngTest.Address = rngMember.Address Then
                GetRangeIndexInRangeArray = lCount
                Exit For
            End If
        Next rngTest
    End If
End Function

Private Function MergedCellsInRange(rngSource As Excel.Range, bRemove As Boolean) As Boolean
'Optionally remove merged cells from a range, copying contents of origin cell back into newly unmerged cells.
'Return if merged cells existed in rngSource.
    Dim rngTemp As Excel.Range
    Dim rngSourceMergedArea As Excel.Range
    Dim rngSourceTemp As Excel.Range
    Dim bFound As Boolean

    For Each rngTemp In rngSource.Cells
        'Set if merged cells exist.
        If bFound = False And (rngTemp.MergeCells) Then
            bFound = True
        End If

        'Remove merged cells
        If bRemove = True And (rngTemp.MergeCells) = True Then
            Set rngSourceMergedArea = rngTemp.MergeArea

            'Get range with source based on value.
            rngTemp.UnMerge 'We assume that address of rngTemp.Address does not change when unmerging.

            'Paste original value back unto newly unmerged cells.
            For Each rngSourceTemp In rngSourceMergedArea.Cells
                If rngSourceTemp.Address <> rngTemp.Address Then
                    rngSourceTemp.Value = rngTemp.Value
'                    Call rngtemp.Copy(rngSourceTemp)
                End If
            Next rngSourceTemp
        End If
    Next rngTemp
    MergedCellsInRange = bFound
End Function

Private Sub DeleteRange(rng As Excel.Range, xlShiftDir As XlDeleteShiftDirection)
'Untested if works on non-contiguous.
'Delete contiguous and non-contiguous ranges, consolidating resultant data.
'.Clear can be used if data doesn't need to be consolidated.
'    Dim vRanges As Variant
'    Dim iCounter As Long

    Application.DisplayAlerts = False
    rng.Delete (xlShiftDir)
    Application.DisplayAlerts = True

' Could modify this to loop through ranges in included range, split is not required.
'    If Not rng Is Nothing Then
'        vRanges = Split(rng.Address, ",")
'
'        'Loop backwards to avoid corrupting data.
'        For iCounter = UBound(vRanges) To 0 Step -1
'            Excel.Range(vRanges(iCounter)).Delete (xlShiftDir)
'        Next iCounter
'    End If
End Sub

Private Sub GetConsolidatedRange(colRange As Collection, rngDestination As Excel.Range, ConsolFunction As Excel.XlConsolidationFunction, Optional bIncludedRowHeader As Boolean = False, Optional bIncludedColHeader As Boolean = False, Optional bCreateLinks As Boolean = False)
'Remember to clear previous range if desired.
'Initially implemented to return a result range that is the average of multiple sets of ranges.
'ConsolFunction = xlUnknown implemented to input first item in collection, skipping others.

'Call like this:
'    Dim colRange As Collection
'    Dim wks As Excel.Worksheet
'    Dim rngDestination As Excel.Range
'    Set colRange = New Collection
'    For Each wks In Application.ActiveWorkbook.Worksheets
'        Call colRange.Add(wks.UsedRange, CStr(wks.Name))
'    Next
'    Call GetConsolidatedRange(colRange, Sheet5.Range("A1"), xlAverage, True, True, False)

    Dim rng As Excel.Range
    Dim strArray() As String
    Dim iCount As Long

    ReDim strArray(colRange.Count - 1) As String

    For Each rng In colRange
        'Cleaning raw inputs.
'        Set rng = RemoveEmptyRowsFromRange(rng) 'Remove extra series, leaving blank columns that will then be filled with 0's.

        On Error Resume Next
        rng.SpecialCells(xlCellTypeBlanks) = 0 'Put 0 in blank cells.  Errors if no values are found.
        On Error GoTo 0 'Turn off error handling and reset err to 0

        strArray(iCount) = rng.Address(, , xlR1C1, True)
'        strArray(iCount) = "'" & VBA.Left(rngDestination.Parent.Parent.FullName, Len(rngDestination.Parent.Parent.Name)) & VBA.Right(rng.Address(, , xlR1C1, True), Len(rng.Address(, , xlR1C1, True)) - 1)
        iCount = iCount + 1
    Next

    'Used as method to input first collection item only (when all values in collection items should be the same)
    'Can include strings
    If ConsolFunction = XlConsolidationFunction.xlUnknown Then
        'These are (or should be) items in model that are exported after final replication only.
        Call colRange(1).Copy
        Call rngDestination.PasteSpecial(xlPasteValues)
    Else
        On Error Resume Next
        Call rngDestination.Consolidate(strArray, ConsolFunction, bIncludedRowHeader, bIncludedColHeader, bCreateLinks)

        If Err.Number = 1004 Then
            Debug.Print "Error in GetConsolidatedRange(): Source reference overlaps destination area."
        End If
    End If
End Sub

Private Function FindInNonContinuousRange(rngLookIn As Excel.Range, strFindValue As String) As Excel.Range
    'Find in non-continuous range. FindInRange() only works on continuous ranges.
    Dim rng As Excel.Range
    For Each rng In rngLookIn
        If rng.Value = strFindValue Then
            Set FindInNonContinuousRange = rng
            Exit For
        End If
    Next
End Function

Private Function FindInRange(rngLookIn As Excel.Range, strFindValue As String) As Excel.Range
' Find strFindValue within rngLookIn and return range.
' only works on continuous ranges.
    Dim rngResult As Excel.Range
    Dim rngFound As Excel.Range
    Dim rngFirstFound As Excel.Range

    Dim rng As Excel.Range
    If rngLookIn(2).Address = rngLookIn(1).Offset(1).Address Then 'Assume Column of cells
        Set rngFirstFound = rngLookIn.Find(strFindValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, False)
    Else 'Assume Row of cells
        Set rngFirstFound = rngLookIn.Find(strFindValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, False)
    End If

    If Not rngFirstFound Is Nothing Then
        Set rngFound = rngFirstFound
        Set rngResult = rngFirstFound

        Set rngFound = rngLookIn.FindNext(rngFound)
        If Not rngFound Is Nothing Then
            Do While rngFound.Address <> rngFirstFound.Address
                Set rngResult = Application.Union(rngResult, rngFound)
                Set rngFound = rngLookIn.FindNext(rngFound)
            Loop
        End If
    End If

errsub:
    Set FindInRange = rngResult
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

Private Function FindDuplicatesInRange(rngFind As Excel.Range, ByVal rngFindArea As Excel.Range, Optional bNotify As Boolean = False) As Excel.Range
'Given range of values passes back range containing two ranges of first found duplicate pair.
    On Error Resume Next
    Dim rng As Excel.Range
    Dim rngFound As Excel.Range
    Dim rngFirstFound As Excel.Range
    Dim rngSourceConstants As Excel.Range
    Dim rngSourceFormulas As Excel.Range
    Dim rngResult As Excel.Range
    Dim strResult As String

    'Get constants and formula values removing blanks from source data set.
    Set rngSourceConstants = rngFindArea.SpecialCells(xlCellTypeConstants) 'Will error if none found
    Set rngSourceFormulas = rngFindArea.SpecialCells(xlCellTypeFormulas) 'Will error if none found

    If Not rngSourceConstants Is Nothing And Not rngSourceFormulas Is Nothing Then
        Set rngFindArea = Application.Union(rngSourceConstants, rngSourceFormulas)
    ElseIf Not rngSourceConstants Is Nothing Then
        Set rngFindArea = rngSourceConstants
    ElseIf Not rngSourceFormulas Is Nothing Then
        Set rngFindArea = rngSourceFormulas
    End If

    On Error GoTo errsub 'Clears previous errors

    For Each rng In rngFind.Cells
'        Application.StatusBar = rng.Row
        Set rngFirstFound = rngFindArea.Find(rng.Value, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows)
        If Not rngFirstFound Is Nothing Then
            If rngFirstFound.Address = rng.Address Then
                Set rngFound = rngFindArea.FindNext(rngFirstFound) 'Always returns address, either original value or new found.
            Else
                Set rngFound = rngFirstFound
            End If

            If Not rngFound Is Nothing Then
                If rngFound.Address <> rng.Address Then
                    Set rngResult = Application.Union(rng, rngFound)
                    GoTo errsub 'Exit Function
                End If
            End If
        End If
    Next

errsub:
    If Err.Number = 0 And bNotify Then
        If Not rngResult Is Nothing Then
            Set FindDuplicatesInRange = rngResult
            Application.EnableEvents = True
            rngResult.Parent.Activate
            rngResult(1).Select
            For Each rng In rngResult
                strResult = strResult & rng.Address(, , , True) & vbCrLf & vbCrLf
            Next

            Call MsgBox("Duplicate Error in:" & vbCrLf & vbCrLf & strResult, vbOKOnly, "Error")
    '        Call MsgBox("Program terminated due to Duplicate Error in:" & vbCrLf & vbCrLf & rng(1).Address(, , , True) & vbCrLf & vbCrLf & "and" & vbCrLf & vbCrLf & rng(rng.Cells.Count).Address(, , , True), vbOKOnly, "Error")
'            Debug.Print rngResult.Areas(1).Address(, , , True) & vbTab & rngResult.Areas(2).Address(, , , True)
        End If
    End If
End Function

Private Function RangeFindLookup(rngTableSource As Excel.Range, strFindValue As String, lOffset As Long, xlExactMatch As Excel.XlLookAt, xlOrient As Orientation) As Excel.Range
    'Function to replace VLookup(columns)/HLookup(rows) functionality.  Searches first column/row of rngTableSource for strFindValue and returns value in offset row/col.  Returns nothing is value not found.
    '(VLookup returns an error if value is not found when looking for exact match).
    'strTemplateName = Application.WorksheetFunction.VLookup(wksSource.Cells(rng.Row, ConvertLetterToColumnNumber(shtCandidateMappings.Range("PSS_CanUniqueID"))).Value, shtCandidateTempLookup.Range("PST_CanTemplateLookupTableNew"), 4, True)
    Dim rngFind As Excel.Range
    rngTableSource.Parent.AutoFilterMode = False 'Turn filter off filters or find can fail.

    If xlOrient = Vertical Then 'Application.WorksheetFunction.VLookup
        Set rngFind = Application.Intersect(rngTableSource, rngTableSource.Cells(1, 1).EntireColumn).Find(strFindValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, False, False)
        If Not rngFind Is Nothing Then
            Set RangeFindLookup = rngFind.Offset(0, lOffset)
        End If
    'Could add this for HLookup
    Else 'xlOrient = Horizontal 'Application.WorksheetFunction.HLookup
        Set rngFind = Application.Intersect(rngTableSource, rngTableSource.Cells(1, 1).EntireRow).Find(strFindValue, , Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlWhole, Excel.XlSearchOrder.xlByColumns, Excel.XlSearchDirection.xlNext, False, False)
        If Not rngFind Is Nothing Then
            Set RangeFindLookup = rngFind.Offset(lOffset, 0)
        End If
    End If
End Function

Private Function FindAll(SearchRange As Excel.Range, FindWhat As Variant, Optional LookIn As XlFindLookIn = Excel.XlFindLookIn.xlValues, Optional LookAt As XlLookAt = Excel.XlLookAt.xlWhole, _
    Optional SearchOrder As XlSearchOrder = Excel.XlSearchOrder.xlByRows, Optional MatchCase As Boolean = False) As Excel.Range
    'http://www.cpearson.com/Excel/RangeFind.htm
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Returns a Range object that contains all cells found.  If none found returns Nothing.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim FoundCell As Excel.Range
    Dim FoundCells As Excel.Range
    Dim LastCell As Excel.Range
    Dim FirstAddr As String

    With SearchRange
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' In order to have Find search for the FindWhat value
        ' starting at the first cell in the SearchRange, we
        ' have to find the last cell in SearchRange and use
        ' that as the cell after which the Find will search.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set LastCell = .Cells(.Cells.Count)
    End With

    ' Do the initial Find.
    Set FoundCell = SearchRange.Find(FindWhat, LastCell, LookIn, LookAt, SearchOrder, MatchCase)
    If Not FoundCell Is Nothing Then
        Set FoundCells = FoundCell
        FirstAddr = FoundCell.Address
        Do
            ' Loop calling FindNext until FoundCell is nothing or we wrap around the first found cell (address is in FirstAddr)
            Set FoundCells = Application.Union(FoundCells, FoundCell)
            Set FoundCell = SearchRange.FindNext(After:=FoundCell)
        Loop Until (FoundCell Is Nothing) Or (FoundCell.Address = FirstAddr)
    End If

    If Not FoundCells Is Nothing Then
'        Set FindAll = Nothing 'Should already be nothing
'    Else
        Set FindAll = FoundCells
    End If
End Function

Private Function FindAllRangesInRange(SearchRange As Excel.Range, FindCell As Excel.Range, Optional LookIn As XlFindLookIn = Excel.XlFindLookIn.xlValues, Optional LookAt As XlLookAt = Excel.XlLookAt.xlWhole, _
    Optional SearchOrder As XlSearchOrder = Excel.XlSearchOrder.xlByRows, Optional MatchCase As Boolean = False) As Excel.Range
    'Find all matching FindCell ranges in SearchRange range, not including FindCell
    'http://www.cpearson.com/Excel/RangeFind.htm
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Returns a Range object that contains all cells found.  If none found returns Nothing.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim FoundCell As Excel.Range
    Dim FoundCells As Excel.Range
    Dim LastCell As Excel.Range
    Dim FirstAddr As String

    Set FindCell = FindCell(1)

    With SearchRange
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' In order to have Find search for the FindWhat value starting at the first cell in the SearchRange, we
        ' have to find the last cell in SearchRange and use that as the cell after which the Find will search.
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set LastCell = .Cells(.Cells.Count)
    End With

    ' Do the initial Find.
    Set FoundCell = SearchRange.Find(FindCell, LastCell, LookIn, LookAt, SearchOrder, MatchCase)
    If Not FoundCell Is Nothing Then
        If FindCell.Address <> FoundCell.Address Then 'Dont include cell we are looking for in results.
            Set FoundCells = FoundCell
            FirstAddr = FoundCell.Address
            Do
                ' Loop calling FindNext until FoundCell is nothing or we wrap around the first found cell (address is in FirstAddr)
                Set FoundCells = Application.Union(FoundCells, FoundCell)
                Set FoundCell = SearchRange.FindNext(After:=FoundCell)
            Loop Until (FoundCell Is Nothing) Or (FoundCell.Address = FirstAddr)
        End If
    End If

    If Not FoundCells Is Nothing Then
'        Set FindAll = Nothing 'Should already be nothing
'    Else
        Set FindAllRangesInRange = FoundCells
    End If
End Function

Private Function FindEmptyCells(rngSource As Excel.Range) As Excel.Range
    On Error GoTo errsub
    Set FindEmptyCells = rngSource.SpecialCells(xlCellTypeBlanks)   'Errors if no values are found.
errsub:
End Function

Private Function FillDownAuto(rngDestination As Excel.Range, Optional FillType As XlAutoFillType = XlAutoFillType.xlFillDefault)
'Autofill first row of rngDestination to rest of area according to XlAutoFillType
    Dim rngHeader As Excel.Range

    If IsEmpty(rngDestination(1).Formula) = False Then 'Check for value or function
        Set rngHeader = rngDestination(1).Resize(, rngDestination.Columns.Count)
    '    Set rngColumn = GetUsedColumn(rngDestination(1)).Resize(, rngDestination.Columns.Count)

        Call rngHeader.AutoFill(rngDestination, FillType)
    End If
End Function

Private Function FillDownValues(rngDestination As Excel.Range)
'Use FillDownAuto()
''Autofill first row of rngDestination to rest of area.
''Copies down values and formulas.  Can be mixed.
'    Dim rngHeader As Excel.Range
'
'    Debug.Assert False 'GBH - Test it ' function below assumes that destination already has values in it.
'    Set rngHeader = rngDestination(1).Resize(, rngDestination.Columns.Count)
''    Set rngColumn = GetUsedColumn(rngDestination(1)).Resize(, rngDestination.Columns.Count)
'
'    Call rngHeader.AutoFill(rngDestination, xlFillCopy)
'
'    If Application.Calculation <> xlCalculationAutomatic Then
'        Application.Calculate
'    End If
End Function

Private Function FillDownFormula(rngDestination As Excel.Range, strFormula As String)
'Use FillDownAuto()
''Autofill contents of first cell set to strFormula, to rest of column.
'    Dim rngColumn As Excel.Range
'
'    Set rngDestination = rngDestination(1)
'    rngDestination.Formula = strFormula
'    Set rngColumn = GetUsedColumnByStartCell(rngDestination)
'
'    Call rngDestination.AutoFill(rngColumn, xlFillCopy)
End Function

Private Function FillDownArea(rngAnchorArea As Excel.Range, rngDestinationArea As Excel.Range, Optional FillType As Excel.XlAutoFillType = xlFillDefault)
'Autofill contents of rngAnchor to rest of rngDestination.
    If Not Application.Intersect(rngAnchorArea, rngDestinationArea) Is Nothing Then
        Call rngAnchorArea.AutoFill(rngDestinationArea, FillType)
    End If
End Function

Private Sub RemoveNumbersStoredAsText(rngSource As Excel.Range)
'Changes "Numbers Stored as Text" to Numbers
'http://www.dailydoseofexcel.com/archives/2006/02/18/number-stored-as-text/
'http://office.microsoft.com/en-us/excel/HP030559001033.aspx

    rngSource.Parent.AutoFilterMode = False 'Turn filter off
    rngSource.Value = rngSource.Value
End Sub

Private Sub ApplyRangeBorder(rng As Excel.Range, AllBorders As Boolean, ThickBorder As Boolean)
    On Error Resume Next

    Dim iBorderInex As Excel.XlBordersIndex

    With rng
        'Outside
        .Borders(xlDiagonalDown).LineStyle = xlNone
        .Borders(xlDiagonalUp).LineStyle = xlNone

        If AllBorders Then
            'Inside
            For iBorderInex = xlEdgeTop To xlInsideHorizontal
                .Borders(iBorderInex).LineStyle = xlContinuous
                .Borders(iBorderInex).Weight = xlThin 'xlMedium
                .Borders(iBorderInex).ColorIndex = vbBlack
            Next
        End If

        If ThickBorder Then
            For iBorderInex = xlEdgeLeft To xlEdgeRight
                .Borders(iBorderInex).LineStyle = xlContinuous
                .Borders(iBorderInex).Weight = xlMedium
                .Borders(iBorderInex).ColorIndex = vbBlack
            Next
        End If

    End With
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Workbook & Worksheet functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub HideAllWorksheets(xlBook As Excel.Workbook)
    Dim wks As Excel.Worksheet

    For Each wks In xlBook.Worksheets   'Don't hide charts.
'        If TypeOf Sh Is Excel.Worksheet Then
        If wks.Visible = xlSheetVisible Then 'Don't modify xlSheetVeryHidden property.
            wks.Visible = xlSheetHidden 'Hide all sheets.
        End If
    Next
End Sub

Private Sub DeleteWorksheet(ByRef sht As Excel.Worksheet)
'Delete sheets suppressing notification.
        Application.DisplayAlerts = False
        sht.Delete
        Application.DisplayAlerts = True
        Set sht = Nothing
End Sub

Private Function GetWorksheetByName(wkbSource As Workbook, strSheetName As String) As Worksheet
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

Private Function GetOpenWorkbookByName(oApplication As Application, strWorkBookName As String) As Excel.Workbook
'Returns workbookobject. Set to nothing if not found.
    On Error Resume Next
    With oApplication
    '    Worksheets (strSheetName)
        Set GetOpenWorkbookByName = .Workbooks(strWorkBookName)
        If Err.Number <> 0 Then
            Set GetOpenWorkbookByName = Nothing
        End If
    End With
    Err.Clear
End Function

Private Function GetWorksheetByCodeName(wkbSource As Workbook, strSheetCodeName As String) As Worksheet
'Used to get worksheet by codename by accessing components in project by name.
'If calling from outside of Excel, can only be used on a remote computer with Excel macro security setting, Trusted Published tab set to "Trust access to Visual Basic Project" on.
'Within Excel can also just use codename.range (for example) to access object by codename.
'Can also be used to determine if worksheet exists.

    On Error Resume Next
    Set GetWorksheetByCodeName = wkbSource.Worksheets(wkbSource.VBProject.VBComponents(strSheetCodeName).Properties("Name").Value)
'        Dim sht As Excel.Worksheet
'        For Each sht In wkbSource.Worksheets
'            If sht.CodeName = strSheetCodeName Then
'                Set GetWorksheetByCodeName = sht
'                Exit For
'            End If
'        Next

    If Err.Number <> 0 Then
        Set GetWorksheetByCodeName = Nothing
    End If
'    With wkbSource
'        Set GetWorksheetByCodeName = .Worksheets(CStr(.VBProject.VBComponents(strSheetCodeName).Properties("Name").value))
'    End With
    Err.Clear
End Function

Private Function GetWorksheetByCodeNameString(wkbSource As Workbook, strSheetCodeName As String) As Excel.Worksheet
'Used to get worksheet object by codename string
    On Error GoTo errsub

    Dim sht As Excel.Worksheet
    For Each sht In wkbSource.Worksheets
        If sht.CodeName = strSheetCodeName Then
            Set GetWorksheetByCodeNameString = sht
            Exit Function
        End If
    Next
errsub:
End Function

Private Sub SetWorkSheetCodeName(wks As Excel.Worksheet, strCodeName As String)
    'Change the codename of a worksheet via code.
    'Could also add check for invalid characters. '(wks.Name, " ", ""), "-", ""), "&", ""), "/", "")
    Dim wksTest As Excel.Worksheet
    Dim bFound As Boolean
    'shtBriggsTable

    'Text is limited to 31 chars
    If Len(strCodeName) > 31 Then
        strCodeName = VBA.Left(strCodeName, 31)
    End If

    'Check if duplicate name first.
    If Not GetWorksheetByCodeName(wks.Parent, strCodeName) Is Nothing Then
        bFound = True
    End If

'    For Each wksTest In wks.Parent.Worksheets
'        If wksTest.CodeName = strCodeName Then
'            bFound = True
'            Exit For
'        End If
'    Next

    If bFound = False Then
        wks.Parent.VBProject.VBComponents(wks.CodeName).Properties("_CodeName") = strCodeName
    Else
        Debug.Print "Duplicate codename found, codename not changed."
    End If
End Sub

Private Function IsWorkbookOpen(strWorkBookName As String) As Boolean
'Return if workbook is currently open.
'Only tests current application instance.
    On Error GoTo errsub
    If Not Application.Workbooks(strWorkBookName) Is Nothing Then
        IsWorkbookOpen = True
    End If
errsub:
End Function

'Possibly use this method instead for WorkbookOpen to test if the file is open rather than looking for the file in the Workbooks collection, which could be in mulitple instances of the application.
'Function FileAlreadyOpen(FullFileName As String) As Boolean
'' returns True if FullFileName is currently in use by another process
'' example: If FileAlreadyOpen("C:\FolderName\FileName.xls") Then...
'    Dim f As Long
'
'    f = FreeFile
'    On Error Resume Next
'    Open FullFileName For Binary Access Read Write Lock Read Write As #f
'    Close #f
'    ' If an error occurs, the document is currently open.
'    If Err.Number <> 0 Then
'        FileAlreadyOpen = True
'        Err.Clear
'        'MsgBox "Error #" & Str(Err.Number) & " - " & Err.Description
'    Else
'        FileAlreadyOpen = False
'    End If
'    On Error GoTo 0 'Turn off error handling
'End Function

Private Function CopyWorksheet(shtOriginal As Excel.Worksheet, strNewName As String, Position As RelativePosition, shtDestinationRefence As Excel.Worksheet, bReplaceDuplicate As Boolean) As Worksheet  'Collection
'Copy existing sheet to new location with new name.
'Added work-around to find added sheet as the name, code name, index can be consistantly identified for reference return.
'MIght be able to use this to get index: wks.Parent.VBProject.VBComponents(wks.CodeName).Properties ("Index")
    Dim sht As Excel.Worksheet
    Dim shtNew As Excel.Worksheet
    Dim shtOldDuplicate As Excel.Worksheet
    Dim wkBook As Excel.Workbook
    Dim colWorksheets As Collection
    Dim bChanged As Boolean
    Set wkBook = shtOriginal.Parent
'    Set CopySheet = Nothing
    Set shtOldDuplicate = GetWorksheetByName(shtOriginal.Parent, strNewName)

    'Delete sheet if it already exists.
    If Not shtOldDuplicate Is Nothing And bReplaceDuplicate Then
        Call DeleteWorksheet(shtOldDuplicate)
    End If

    'If origin sheet exists and new sheet name does not exist.
    If Not shtOriginal Is Nothing And shtOldDuplicate Is Nothing Then
        'Work-around for limitation in copy method in that it doesn't return reference to new object created with copy method.
        'Keep track of original worksheets
        Set colWorksheets = New Collection
        For Each sht In shtOriginal.Parent.Worksheets
            colWorksheets.Add sht, sht.CodeName
        Next

        'Can't copy VeryHidden sheets, so temporarily change.
        If shtOriginal.Visible = xlSheetVeryHidden Then
            bChanged = True
            shtOriginal.Visible = xlSheetHidden
        End If

        If Position = After Then
            Call shtOriginal.Copy(, shtDestinationRefence)
        Else
            Call shtOriginal.Copy(shtDestinationRefence)
        End If

        'Put back property.
        If bChanged = True Then
            shtOriginal.Visible = xlSheetVeryHidden
        End If

        'Find added worksheet by comparing with original collection.
        On Error Resume Next
        For Each sht In shtOriginal.Parent.Worksheets
            If colWorksheets(sht.CodeName) Is Nothing Then 'Look for worksheet by key name
                Set shtNew = sht 'Falls through when error occurs (value not found in collection)
                Exit For
            End If
        Next
        On Error GoTo 0

        'Check that sheet name doesn't already exist
        If (strNewName) <> vbNullString Then
            shtNew.name = strNewName
        End If
        Set CopyWorksheet = shtNew
    ElseIf Not shtOldDuplicate Is Nothing Then '
        Call MsgBox("Sheet named: " & strNewName & vbCrLf & _
        "already exists in sheet: " & shtOriginal.Parent.name & vbCrLf & _
        "in folder: " & shtOriginal.Parent.Path & vbCrLf & vbCrLf & _
        "This sheet must be removed before continuing.", vbOKOnly, "Copy Error")
    End If

    Exit Function
End Function

Private Sub MoveWorksheetToFirstTab(wkbTarget As Excel.Workbook, wksTarget As Excel.Worksheet)
    If wkbTarget.Worksheets.Count > 0 Then
        If Not wksTarget Is wkbTarget.Worksheets(1) Then
            Call wksTarget.Move(, wkbTarget.Worksheets(wkbTarget.Worksheets.Count))
        End If
    End If
End Sub

Private Sub MoveWorksheetToLastTab(wkbTarget As Excel.Workbook, wksTarget As Excel.Worksheet)
    If wkbTarget.Worksheets.Count > 0 Then
        If Not wksTarget Is wkbTarget.Worksheets(wkbTarget.Worksheets.Count) Then
            Call wksTarget.Move(, wkbTarget.Worksheets(wkbTarget.Worksheets.Count))
        End If
    End If
End Sub

Private Sub SortTable(Table As listObject, rngToSort As Excel.Range, SortOrder As Excel.XlSortOrder)
    Table.Sort.SortFields.Clear
    Table.Sort.SortFields.Add rngToSort, xlSortOnValues, SortOrder, xlSortNormal
    With Table.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

Private Sub SortRange(rng As Excel.Range)
    Call rng.Sort(rng)
End Sub

Private Function SortRangeRowsAscending(rngToSort As Excel.Range, RelativeColumnToSortBy As Long, Optional HasHeader As Excel.XlYesNoGuess = Excel.XlYesNoGuess.xlYes)
    'RelativeColumnToSortBy is relative to rngToSort
    Dim rng As Excel.Range

    rngToSort.Parent.Sort.SortFields.Clear

    Set rng = rngToSort.Parent.Cells(rngToSort(1).Row, rngToSort.Columns(RelativeColumnToSortBy).Column)
    Call rngToSort.Sort(rng, xlAscending, , , , , , HasHeader, 1, False, xlTopToBottom, , xlSortTextAsNumbers)
End Function

Private Function SortRangeColumnsAscending(rngToSort As Excel.Range, RelativeRowToSortBy As Long, Optional HasHeader As Excel.XlYesNoGuess = Excel.XlYesNoGuess.xlYes)
    'RelativeRowToSortBy is relative to rngToSort
    Dim rng As Excel.Range

    rngToSort.Parent.Sort.SortFields.Clear

    Set rng = rngToSort.Parent.Cells(rngToSort.Rows(RelativeRowToSortBy).Row, rngToSort(1).Column)
    Call rngToSort.Sort(rng, xlAscending, , , , , , HasHeader, 1, False, xlLeftToRight, , xlSortTextAsNumbers)
End Function

Private Function SortAllWorksheets(wkbTarget As Excel.Workbook, ByRef ErrorText As String, Optional ByVal SortDescending As Boolean = False) As Boolean
    SortAllWorksheets = SortWorksheetsByName(wkbTarget, 0, 0, ErrorText, SortDescending)
End Function

Private Function SortWorksheetsByName(wkbTarget As Excel.Workbook, ByVal FirstToSort As Long, ByVal LastToSort As Long, ByRef ErrorText As String, Optional ByVal SortDescending As Boolean = False) As Boolean
    'http://www.cpearson.com/excel/sortws.aspx 'Modified

    ' This sorts the worskheets from FirstToSort to LastToSort by name in either ascending (default) or descending order.
    ' If successful, ErrorText is vbNullString and the function returns True. If unsuccessful, ErrorText gets the reason why the function failed and the function returns False.
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'FirstToSort is the index (position) of the first worksheet to sort.
    'LastToSort is the index (position) of the laset worksheet to sort. If either both FirstToSort and LastToSort are less than or equal to 0, all sheets in the workbook are sorted.
    'ErrorText is a variable that will receive the text description of any error that may occur.
    'SortDescending is an optional parameter to indicate that the sheets should be sorted in descending order. If True, the sort is in descending order. If omitted or False, the sort is in ascending order.

    Dim m As Long
    Dim n As Long
'    Dim WB As Excel.Workbook
    Dim b As Boolean

'    Set WB = wkbTarget 'Worksheets.Parent
    ErrorText = vbNullString

    If wkbTarget.ProtectStructure = True Then
        ErrorText = "Workbook is protected."
        SortWorksheetsByName = False
    End If

    ' If First and Last are both 0, sort all sheets.
    If (FirstToSort = 0) And (LastToSort = 0) Then
        FirstToSort = 1
        LastToSort = wkbTarget.Worksheets.Count
    Else
        ' More than one sheet selected. We can sort only if the selected sheet are adjacent.
        b = TestFirstLastSort(FirstToSort, LastToSort, ErrorText)
        If b = False Then
            SortWorksheetsByName = False
            Exit Function
        End If
    End If

    ' Do the sort, essentially a Bubble Sort.
    For m = FirstToSort To LastToSort
        For n = m To LastToSort
            If SortDescending = True Then
                If StrComp(wkbTarget.Worksheets(n).name, wkbTarget.Worksheets(m).name, vbTextCompare) > 0 Then
                    Call wkbTarget.Worksheets(n).Move(wkbTarget.Worksheets(m))
                End If
            Else
                If StrComp(wkbTarget.Worksheets(n).name, wkbTarget.Worksheets(m).name, vbTextCompare) < 0 Then
                    Call wkbTarget.Worksheets(n).Move(wkbTarget.Worksheets(m))
                End If
            End If
        Next n
    Next m

    SortWorksheetsByName = True
End Function

Private Function TestFirstLastSort(FirstToSort As Long, LastToSort As Long, ByRef ErrorText As String) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' This ensures FirstToSort and LastToSort are valid values. If successful returns True and sets ErrorText to vbNullString. If unsuccessful, returns
    ' False and set ErrorText to the reason for failure.
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ErrorText = vbNullString
    If FirstToSort <= 0 Then
        TestFirstLastSort = False
        ErrorText = "FirstToSort is less than or equal to 0."
        Exit Function
    End If

    If FirstToSort > Worksheets.Count Then
        TestFirstLastSort = False
        ErrorText = "FirstToSort is greater than number of sheets."
        Exit Function
    End If

    If LastToSort <= 0 Then
        TestFirstLastSort = False
        ErrorText = "LastToSort is less than or equal to 0."
        Exit Function
    End If

    If LastToSort > Worksheets.Count Then
        TestFirstLastSort = False
        ErrorText = "LastToSort greater than number of sheets."
        Exit Function
    End If

    If FirstToSort > LastToSort Then
        TestFirstLastSort = False
        ErrorText = "FirstToSort is greater than LastToSort."
        Exit Function
    End If

    TestFirstLastSort = True
End Function

Private Function DayofYear(oDate As Date) As Long
    DayofYear = DateDiff("d", DateSerial(Year(oDate), 1, 1), oDate)
End Function

Private Function DaysPerMonth(oDate As Date) As Variant
'Returns the number of days in a month.
'Defines the 0th day of next month which is the same as the last day of this month.
    DaysPerMonth = Day(DateSerial(Year(oDate), Month(oDate) + 1, 0))
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Named Ranges/Formulas
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'http://www.cpearson.com/excel/named.htm

Private Function SetNamedRange(NamedRange As String, ReferesTo As Variant, Visible As Boolean, Optional Comment As String) As Boolean
    'Create Named Range, ReferesTo can be range, string formula, etc.
    'optionally set visibility
    On Error Resume Next 'Will error if it doesn't exist.
    Call ThisWorkbook.Names(NamedRange).Delete
    On Error GoTo 0

    With ThisWorkbook.Names.Add(NamedRange, ReferesTo, Visible)
        .Comment = Comment
    End With
    SetNamedRange = True
End Function

Private Function NamedRangeExists(sNamedRangeName As String) As Boolean
'http://www.j-walk.com/ss/excel/tips/tip54.htm
'   Returns TRUE if the range name exists
    Dim n As name
'    RangeNameExists = False
    For Each n In ActiveWorkbook.Names
        If VBA.Strings.UCase$(n.name) = VBA.Strings.UCase$(sNamedRangeName) Then
            NamedRangeExists = True
            Exit Function
        End If
    Next n
End Function

Private Sub CreateExcelSheetLevelNames(wks As Excel.Worksheet, lHeaderOffset As Long, xlOrient As Excel.XlRowCol)
'For automation of the creation of Name Formulas at the sheet level.
'Once called, sheet rows and columns can then be accessed by name which provide values in a dynamic fashion (for chart reference values).
'Will create based on usedrange which includes unseen cell items.
'Automatically replaces existing values of the same name.
    Dim rng As Excel.Range
    Dim rngUsed As Excel.Range
'    Dim strTest As String

    If xlOrient = xlRows Then 'Rows
        Set rngUsed = Application.Intersect(wks.UsedRange, wks.Range("$A$1").EntireColumn)
    Else
        Set rngUsed = Application.Intersect(wks.UsedRange, wks.Range("$A$1").EntireRow)
    End If

    For Each rng In rngUsed
        If xlOrient = xlRows Then 'Rows
'            strTest = "=OFFSET('" & wks.Name & "'!" & rng.Address(True, True, xlR1C1) & ",0," & lHeaderOffset & ",1,COUNTA('" & wks.Name & "'!" & rng.EntireRow.Address(True, True, xlR1C1) & ")-" & lHeaderOffset & ")"
'            Debug.Print strTest
            wks.Names.Add "Row" & CStr(rng.Row), , , , , , , , , "=OFFSET('" & wks.name & "'!" & rng.Address(True, True, xlR1C1) & ",0," & lHeaderOffset & ",1, COUNTA('" & wks.name & "'!" & rng.EntireRow.Address(True, True, xlR1C1) & ")-" & lHeaderOffset & ")"
            wks.Names("Row" & CStr(rng.Row)).Comment = "Auto-generated using CreateExcelSheetLevelName()"
        Else 'Columns
'            strTest = "=OFFSET('" & wks.Name & "'!" & rng.Address(True, True, xlR1C1) & "," & lHeaderOffset & ",0,COUNTA('" & wks.Name & "'!" & rng.EntireColumn.Address(True, True, xlR1C1) & ")-" & lHeaderOffset & ",1)"
'            Debug.Print strTest
            wks.Names.Add "Column" & GetRangeColumnLetter(rng), , , , , , , , , "=OFFSET('" & wks.name & "'!" & rng.Address(True, True, xlR1C1) & "," & lHeaderOffset & ",0,COUNTA('" & wks.name & "'!" & rng.EntireColumn.Address(True, True, xlR1C1) & ")-" & lHeaderOffset & " ,1)"
            wks.Names("Column" & GetRangeColumnLetter(rng)).Comment = "Auto-generated using CreateExcelSheetLevelName()"
        End If
    Next
End Sub

Private Sub SetExcelChartSeriesToNamedFormulas(xlsChart As Excel.Chart, wksData As Excel.Worksheet, lHeaderOffset As Long, xlOrient As Excel.XlRowCol, iSeriesCount As Long)
'Untested, containes hard coded values.
'Example call:
'Call SetExcelChartSeriesToNamedFormulas(Sheet2.ChartObjects(1).Chart, Sheet2, 1, xlColumns)
'Set source data values in chart with named function values.
    Dim obj As Excel.Chart
    Dim rng As Excel.Range
    Dim rngUsed As Excel.Range
    Dim strTest As String
    Dim iCount As Long: iCount = 1
'    Dim iSeriesCount As Long: iSeriesCount = 13

    If xlOrient = xlRows Then 'Rows
        Set rngUsed = Application.Intersect(wksData.UsedRange, wksData.Range("$A$1").EntireColumn)
    Else
        Set rngUsed = Application.Intersect(wksData.UsedRange, wksData.Range("$A$1").EntireRow)
    End If

    'Remove all previous series adding new ones according to count
    Do Until xlsChart.SeriesCollection.Count = 0
        xlsChart.SeriesCollection(1).Delete
    Loop

'    'Add series
'    For iCount = 1 To iSeriesCount
'        With xlsChart.SeriesCollection.NewSeries
''            .Values = rngChtXVal.Offset(, iColumn - 1)
''            .XValues = rngChtXVal
''            .Name = rngChtData(1, iColumn)
'        End With
'    Next

    iCount = 1
    For iCount = 1 To iSeriesCount
'    For Each rng In rngUsed
        If xlOrient = xlRows Then 'Rows
            With xlsChart.SeriesCollection.NewSeries 'Add series
    '            strTest = "='Summary | Avg Scenarios'!Row" & CStr(rng.Row)
    '            Debug.Print strTest
                .Values = "='Summary | Avg Scenarios'!Row" & CStr(iCount + lHeaderOffset)
    '            xlsChart.SeriesCollection(iCount).Values = "='Summary | Avg Scenarios'!Row" & CStr(rng.Row)
            End With
        Else 'Columns
            With xlsChart.SeriesCollection.NewSeries 'Add series
    '            strTest = "='Summary | Avg Scenarios'!Column" & GetRangeColumnLetter(rng)
    '            Debug.Print strTest
                .name = "='Summary | Avg Scenarios'!$" & GetRangeColumnLetter(wksData.Cells(, iCount + lHeaderOffset)) & "$1"
                .Values = "='Summary | Avg Scenarios'!Column" & GetRangeColumnLetter(wksData.Cells(, iCount + lHeaderOffset))
    '            xlsChart.SeriesCollection(iCount).Values = "='Summary | Avg Scenarios'!Column" & GetRangeColumnLetter(rng)
                .XValues = "='Summary | Avg Scenarios'!$A$2"
            End With
        End If
    Next

'    iCount = 1
'    For Each rng In rngUsed
'        If xlOrient = xlRows Then 'Rows
''            strTest = "='Summary | Avg Scenarios'!Row" & CStr(rng.Row)
''            Debug.Print strTest
'            xlsChart.SeriesCollection(iCount).Values = "='Summary | Avg Scenarios'!Row" & CStr(rng.Row)
'        Else 'Columns
''            strTest = "='Summary | Avg Scenarios'!Column" & GetRangeColumnLetter(rng)
''            Debug.Print strTest
'            xlsChart.SeriesCollection(iCount).Values = "='Summary | Avg Scenarios'!Column" & GetRangeColumnLetter(rng)
'        End If
'        iCount = iCount + 1
'    Next

'    Set obj = wks.ChartObjects("Chart 1").Chart '.Activate
'    obj.SeriesCollection(1).Values = "='Summary | Avg Scenarios'!ColumnB"
'    obj.SeriesCollection(2).Values = "='Summary | Avg Scenarios'!ColumnC"
'    obj.SeriesCollection(3).Values = "='Summary | Avg Scenarios'!ColumnD"
'    obj.SeriesCollection(4).Values = "='Summary | Avg Scenarios'!ColumnE"
'    obj.SeriesCollection(5).Values = "='Summary | Avg Scenarios'!ColumnF"
'    obj.SeriesCollection(6).Values = "='Summary | Avg Scenarios'!ColumnG"
'    obj.SeriesCollection(7).Values = "='Summary | Avg Scenarios'!ColumnH"
'    obj.SeriesCollection(8).Values = "='Summary | Avg Scenarios'!ColumnI"
'    obj.SeriesCollection(9).Values = "='Summary | Avg Scenarios'!ColumnJ"
'    obj.SeriesCollection(10).Values = "='Summary | Avg Scenarios'!ColumnK"
'    obj.SeriesCollection(11).Values = "='Summary | Avg Scenarios'!ColumnL"
'    obj.SeriesCollection(12).Values = "='Summary | Avg Scenarios'!ColumnM"
'    obj.SeriesCollection(13).Values = "='Summary | Avg Scenarios'!ColumnN"
'    obj.SeriesCollection(14).Values = "='Summary | Avg Scenarios'!ColumnO"
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' File/Directory/Path Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetXLSFileValue(strFullFilePath As String, strSheet As String, strCell As String) As String
'Tested - See p.354 & 894 in "Excel 2003 Power Programming" - John Walkenbach
'Returns a value from a CLOSED workbook (doesn't open the workbook to get the value)
'Can loop on this function to get multiple values.
    Dim arg As String
    Dim ret As Variant
    'Check if file exists.
    If VBA.Dir(strFullFilePath) <> vbNullString Then
        'Get range
        arg = "'" & ParsePath(strFullFilePath, Path) & "[" & ParsePath(strFullFilePath, FileName) & "]" & strSheet & "'!" & Application.Range(strCell).Address(, , xlR1C1)
'        Debug.Print arg

        ret = Application.ExecuteExcel4Macro(arg)

        If Not IsError(ret) And Not IsEmpty(ret) Then
            GetXLSFileValue = ExecuteExcel4Macro(arg)
        End If
    End If
End Function

Private Function BrowseForFile(strFileFilter As String, lngIndex As Long, strDialogTitle As String, strButtonText As String) As Variant
'Included function call "GetOpenFilename" only exists in Excel.  So only use with Excel.  Otherwise use modCommonDlg.
'Browse for file (which must exist). Similar to Common Dialog Control in VB6 (comdlg32.ocx)
'Example call: BrowseForFile("Microsoft Office Excel File (*.xls), *.xls", 1, "Select File", "Select")
    Dim FileToOpen As Variant

    FileToOpen = Application.GetOpenFileName(strFileFilter, lngIndex, strDialogTitle, strButtonText)

    'If Cancel was not selected.
    If FileToOpen <> vbFalse Then
        BrowseForFile = FileToOpen
    Else
        Call MsgBox("File not found.", vbInformation, "File Error")
    End If

'    'Could also return file/folder object using:
'    Dim objFSO As Object
'    Dim fFileCurrent As File
'    Dim fFolderCurrent As File

'    Set objFSO = CreateObject("Scripting.FileSystemObject") 'New FileSystemObject
'    Set fFileCurrent = objFSO.GetFile(FileToOpen)
'    Set fFolderCurrent = objFSO.GetFolder(FileToOpen)
End Function

Private Sub LoadFile(strFilePath As String, Optional appStyle As VbAppWinStyle = VbAppWinStyle.vbNormalFocus)
'Loads a file with the default provider.  Can also load applications.
    On Error GoTo errsub

    Call Shell(strFilePath, appStyle)   'vbMinimizedFocus)
errsub:
End Sub

Private Function CloseFile(strFilePath As String, Optional SaveBeforeClose As Boolean = False) As Boolean
    On Error GoTo errsub
    Dim FilePtr As Object

    If IsFileOpen(strFilePath) = True Then
        Set FilePtr = GetObject(strFilePath)
'        Application.DisplayAlerts = False   ' TURN OFF EXCEL DISPLAY
        If SaveBeforeClose = True Then
            FilePtr.Save
        End If
        FilePtr.Close
    End If
    CloseFile = True
errsub:
    Set FilePtr = Nothing
'    Application.DisplayAlerts = True
End Function

Private Function IsFileOpen(strFullFilePath As String) As Boolean
'http://www.xcelfiles.com/IsFileOpen.html
'IsFileOpen() only returns true if open in Excel (or possibly other DB files)
    Dim hdlFile As Long

    'Error is generated if you try opening a File for ReadWrite lock >> MUST BE OPEN!
    If VBA.Dir(strFullFilePath) <> vbNullString Then
        On Error GoTo errsub:
        hdlFile = FreeFile
'        Open strFullFilePath For Random Access Read Write Lock Read Write As #hdlFile 'Works for Excel files, but not text files.
        Open strFullFilePath For Input Lock Read As #hdlFile  'http://www.cpearson.com/excel/IsFileOpen.aspx
        'Open strFullFilePath For Random Access Read Write Lock Read Write As hdlFile
        
        Close hdlFile
    End If

    IsFileOpen = False
    Exit Function

errsub: 'Someone has file open
    IsFileOpen = True
'    Close hdlFile
End Function

Private Function IsExcelOpen() As Boolean
' Procedure dectects a running Excel and registers it in the Running Object Table.
'From GetObject function in the Help file.
    Dim hWnd As Long
' If Excel is running this API call returns its handle.
    hWnd = FindWindow("XLMAIN", vbNullString)
    If hWnd = 0 Then    ' 0 means Excel not running.
        IsExcelOpen = False
    Else
    ' Excel is running so enter it in the Running Object Table.
        Call RegisterInROT(hWnd)
        IsExcelOpen = True
    End If
End Function

Private Function RegisterInROT(hWnd As Long)
    'Use the SendMessage API function to enter Application into Running Object Table.
    'If multiple instances are running, only the first launched is entered, so enter in ROT.
    Call SendMessage(hWnd, WM_USER + 18, 0, 0)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Control/Object/Collection Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetFormControlObjectsCollectionByType(sht As Object, tObjectType As Excel.XlFormControl) As Collection
'sht can be worksheet or chart.
'http://www.rondebruin.nl/controlsobjectsworksheet.htm

'MSForms.CommandButton
'MSForms.CheckBox
'MSForms.TextBox
'MSForms.OptionButton
'MSForms.ListBox
'MSForms.ComboBox
'MSForms.ToggleButton
'MSForms.SpinButton
'MSForms.ScrollBar
'MSForms.Label
'MSForms.Image

    Dim col As Collection
    Dim obj As Object

    Set col = New Collection

    Select Case tObjectType
        Case XlFormControl.xlButtonControl
            For Each obj In sht.Buttons
                col.Add obj, obj.name
            Next
        Case XlFormControl.xlCheckBox
            For Each obj In sht.CheckBoxes
                col.Add obj, obj.name
            Next
        Case XlFormControl.xlDropDown
            For Each obj In sht.DropDowns
                col.Add obj, obj.name
            Next
'        Case XlFormControl.xlEditBox
'            For Each obj In sht
'                col.Add obj, obj.Name
'            Next
        Case XlFormControl.xlGroupBox
            For Each obj In sht.GroupBoxes
                col.Add obj, obj.name
            Next
        Case XlFormControl.xlLabel
            For Each obj In sht.Labels
                col.Add obj, obj.name
            Next
        Case XlFormControl.xlListBox
            For Each obj In sht.ListBoxes
                col.Add obj, obj.name
            Next
        Case XlFormControl.xlOptionButton '(Radio Button)
            For Each obj In sht.OptionButtons
                col.Add obj, obj.name
            Next
        Case XlFormControl.xlScrollBar
            For Each obj In sht.ScrollBars
                col.Add obj, obj.name
            Next
        Case XlFormControl.xlSpinner
            For Each obj In sht.Spinners
                col.Add obj, obj.name
            Next
    End Select

    Set GetFormControlObjectsCollectionByType = col
End Function

Private Function GetSelectedObject(wsActiveWorkSheet As Worksheet, tObjectType As XlFormControl) As Object
'Returns object of type specified.
'Implemented for Option Button(Radio buttons) created from Forms Toolbar in VBA
    Select Case tObjectType
        Case xlOptionButton
            Dim ctlTest As OptionButton

            For Each ctlTest In wsActiveWorkSheet.OptionButtons
            'For Each ctlTest In ActiveSheet.OptionButtons
                If ctlTest.Value = Excel.Constants.xlOn Then 'The selected one, others are Excel.Constants.xlOff -4146
                    Set GetSelectedObject = ctlTest
                    Exit For
                End If
            Next ctlTest
        Case Else
    End Select
    'TypeName(GetSelectedObject) =
End Function

Private Function SetSelectedObject(wsActiveWorkSheet As Worksheet, tObjectType As XlFormControl, strCaption As String) As Object
'Selects object of type specified.
'Implemented for Option Button(Radio buttons) created from Forms Toolbar in VBA
    Select Case tObjectType
        Case xlOptionButton
            Dim ctlTest As OptionButton

            For Each ctlTest In wsActiveWorkSheet.OptionButtons
            'For Each ctlTest In ActiveSheet.OptionButtons
                If ctlTest.Caption = strCaption Then
                    Set SetSelectedObject = ctlTest
                    ctlTest.Value = Excel.Constants.xlOn '1
                ElseIf ctlTest.Value = 1 Then 'If selected
                    ctlTest.Value = Excel.Constants.xlOff '-4146 'unselect
                End If
            Next ctlTest
        Case Else
    End Select
    'TypeName(SetSelectedObject) =
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' General Utility Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCollectionIndexByName(sht As Excel.Worksheet, strName As String) As Long
'Return index into collection given name.
    Dim objProp As CustomProperty
    Dim lCount As Long
    lCount = 1
    For Each objProp In sht.CustomProperties
        If objProp.name = strName Then
            GetCollectionIndexByName = lCount
            Exit For
        End If
        lCount = lCount + 1
    Next
End Function

Private Function IsInContainer(wkb As Excel.Workbook) As Boolean
    'Test whether workbook is in a container such as IE.
    'http://www.ozgrid.com/forum/showthread.php?t=28842
    Dim objTemp As Object
    On Error Resume Next

    Set objTemp = wkb.Container
    If Not objTemp Is Nothing Then
        IsInContainer = True
    Else
        IsInContainer = False
    End If
End Function

Private Function GetExcelCaller() As String
    Dim strResult As String
    Select Case TypeName(Application.Caller)
        Case "Range"
            strResult = Application.Caller.Address
        Case "String"
            strResult = Application.Caller
        Case "Error"
            strResult = "Error"
        Case Else
            strResult = "unknown"
    End Select
    GetExcelCaller = strResult
'    MsgBox "caller = " & strResult
End Function

Private Sub SetWorksheetsTabColor(colSheets As Collection, lvbColor As Long)
'Set visibility property of passed in colleciton of sheets.
    Dim sht As Excel.Worksheet

    For Each sht In colSheets
        With sht
            .Tab.Color = lvbColor
            .Tab.TintAndShade = 0
        End With
    Next
End Sub

Private Sub SetWorksheetsVisibleProperty(colSheets As Collection, XlVisibleProperty As Excel.XlSheetVisibility)
'Set visibility property of passed in colleciton of sheets.
    Dim sht As Excel.Worksheet

    For Each sht In colSheets
        sht.Visible = XlVisibleProperty 'xlSheetHidden ' xlSheetVeryHidden 'xlSheetVisible
    Next
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Custom Property Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetWorksheetProperty(wksActive As Excel.Worksheet, strPropertyName As String) As Variant
    Dim objProp As CustomProperty

    Set objProp = GetWorksheetPropertyObject(wksActive, strPropertyName)
    If Not objProp Is Nothing Then
        GetWorksheetProperty = objProp.Value
    End If
End Function

Private Sub SetWorksheetProperty(wksActive As Excel.Worksheet, strPropertyName As String, vPropertyValue As Variant)
    Dim objProp As CustomProperty

    Set objProp = GetWorksheetPropertyObject(wksActive, strPropertyName)
    If objProp Is Nothing Then 'Add it
        Call wksActive.CustomProperties.Add(strPropertyName, vPropertyValue)
    Else 'Update It
        objProp.Value = vPropertyValue
    End If
End Sub

Private Function GetWorksheetPropertyObject(wksActive As Excel.Worksheet, strPropertyName As String) As CustomProperty
'    Set vProps = shtActive.CustomProperties(strPropertyName) 'Doesn't return object so iterate.
    Dim objProp As CustomProperty 'Object

    For Each objProp In wksActive.CustomProperties
        If objProp.name = strPropertyName Then
            Set GetWorksheetPropertyObject = objProp
            Exit For
        End If
    Next
End Function

Private Function GetWorkbookProperty(wkbActive As Excel.Workbook, strPropertyName As String, Optional bIsCustomProperty As Boolean = True) As Variant
'http://www.cpearson.com/excel/docprop.htm 'Modified/Simplified
'Check for empty return with IsEmpty().
    On Error Resume Next

    If bIsCustomProperty = True Then
        GetWorkbookProperty = wkbActive.CustomDocumentProperties(strPropertyName).Value
    Else
        GetWorkbookProperty = wkbActive.BuiltinDocumentProperties(strPropertyName).Value
    End If

End Function

Private Sub SetWorkbookProperty(wkbActive As Excel.Workbook, strPropertyName As String, vPropertyValue As Variant, Optional bIsCustomProperty As Boolean = True)
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

Private Function TerminateExcel()
    'Untested
    'http://www.thescripts.com/forum/thread368443.html
    'Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal
    'lpWindowName As String) As Int32
    'Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Int32, ByVal wMsg
    'As Int32, ByVal wParam As Int32, ByVal lParam As Int32) As Int32
'    Dim ClassName As String
'    Dim WindowHandle As Long ' Int32
'    Dim ReturnVal As Long 'Int32
'    Const WM_QUIT = &H12
'
'    Do
'
'    ClassName = "XLMain"
'    WindowHandle = FindWindow(ClassName, Nothing)
'
'    If WindowHandle Then
'        ReturnVal = PostMessage(WindowHandle, WM_QUIT, 0, 0)
'    End If
'
'    Loop Until WindowHandle = 0

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Outline Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function RemoveAllOutlines(wks As Excel.Worksheet)
    'Removes all outline levels from worksheet.
    wks.UsedRange.RemoveSubtotal
End Function

Private Function LoaderApplyOutlineRowLevels()
    Call ApplyOutlineRowLevels(ActiveSheet.Columns(1))
End Function

Private Function ApplyOutlineRowLevels(rngLevels As Excel.Range)
    'Removes all outline levels then applies new levels to rows of entire sheet based on values in rngLevels.
    'Levels can be appled with keyboard Shift+ Alt + right arrow | left arrow.
    'Outline levels are 1 based.
    Dim rng As Excel.Range

    'Remove previous subtotal
    Call RemoveAllOutlines(rngLevels.Parent)

    Set rngLevels = Intersect(rngLevels(1).EntireColumn, rngLevels.Parent.UsedRange)
    For Each rng In rngLevels
        If rng.Value > 0 And IsNumeric(rng.Value) Then 'could be summary row.
            rng.EntireRow.OutlineLevel = val(rng.Value)
        End If
    Next
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Validation Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
 Private Sub AddValidationCirclesForPrinting()
 'Could be modified to circle cells that are found to contain incorrect values in validation routines.
 'http://support.microsoft.com/default.aspx?scid=kb;en-us;213773
    Dim DataRange As Excel.Range
    Dim C As Excel.Range
    Dim Count As Long
    Dim o As Shape

    'Set an object variable to all of the cells on the active
    'sheet that have data validation -- if an error occurs, run
    'the error handler and end the procedure
    On Error GoTo errhandler
    Set DataRange = Cells.SpecialCells(xlCellTypeAllValidation)
    On Error GoTo 0

    Count = 0

    'Loop through each cell that has data validation
    For Each C In DataRange
       'If the validation value for the cell is false, then draw a circle around the cell. Set the circle's fill to
       'invisible, the line color to red and the line weight to 1.25
       If Not C.Validation.Value Then
           Set o = ActiveSheet.Shapes.AddShape(msoShapeOval, C.Left - 2, C.Top - 2, C.Width + 4, C.Height + 4)
           o.Fill.Visible = msoFalse
           o.Line.ForeColor.SchemeColor = 10
           o.Line.Weight = 1.25

           'Change the name of the shape to InvalidData_ + count
           Count = Count + 1
           o.name = "InvalidData_" & Count
       End If
   Next
   Exit Sub

errhandler:
   MsgBox "There are no cells with data validation on this sheet."
 End Sub

 Private Sub RemoveValidationCircles()
 'http://support.microsoft.com/default.aspx?scid=kb;en-us;213773
    Dim shp As Shape

    'Remove each shape on the active sheet that has a name starting with InvalidData_
    For Each shp In ActiveSheet.Shapes
       If shp.name Like "InvalidData_*" Then shp.Delete
    Next
 End Sub

'Add hyperlink to cell.
' - untested - here for reference.
Private Sub AddHyperlink(rngAnchor As Excel.Range, rngLink As Excel.Range)
    Call rngAnchor.Parent.Hyperlinks.Add(rngLink, "", rngAnchor.Address(, , , True), "Goto: " & rngLink.Value, rngLink.Value)
End Sub

Private Sub InsertComment(rng As Excel.Range, strComment As String)
    rng.ClearComments
    rng.AddComment (VBA.Trim(strComment))
    rng.Comment.Shape.TextFrame.AutoSize = True
End Sub

Private Function ModifyValidationAddress(rngValidation As Excel.Range, rngTarget As Excel.Range) As Boolean
    'Make sure to include entire range if rngValidation is merged cell.

    Call rngValidation.Validation.Modify(, , , "=" & rngTarget.Address)
    ModifyValidationAddress = True
End Function

Private Function SetInCellDropDowns(rngDestination As Excel.Range, vValues As Variant)
'Pass in single dimension array values (Possibly created with RangeArrayToArray()) or array
'If longer is needed instead pass in Named Range in the format "=<Named Range>"
'Length of vValues is limited to 255 characters by Excel.
'Can't set if sheet protection is on.
    Dim strValidation As String

    With rngDestination
'        .Validation.Delete 'Remove previous settings
'        rng.Parent.Activate 'Has to have focus or errors.
        If varType(vValues) = vbArray Or varType(vValues) = 8204 Then '8204 = Variant array
            strValidation = VBA.Trim(Join(vValues, ","))
        ElseIf varType(vValues) = vbString Then
            strValidation = VBA.Trim(vValues)
        Else
            Call MsgBox("Unknown paramater type.", vbInformation, "Internal Error")
        End If

        If Len(strValidation) > 255 Then
            Call MsgBox("Validation not added because the length of the values list exceeds the limits of Excel.", vbInformation, "Internal Error")
        Else
            Call .Validation.Add(xlValidateList, xlValidAlertStop, xlBetween, strValidation)
        End If
    End With
End Function

Private Function GetInCellValidationSelectedIndex(rngCell As Excel.Range) As Long
'Return 0 based index of selected item in Data Validation List marked as In-cell dropdown.
    Dim vArray As Variant
    Dim i As Long

    Set rngCell = rngCell(1)
    GetInCellValidationSelectedIndex = -1

    If rngCell.Validation.Formula1 <> vbNullString Then
        vArray = Split(rngCell.Validation.Formula1, ",")
        For i = 0 To UBound(vArray)
            If VBA.Trim(rngCell.Value) = VBA.Trim(vArray(i)) Then
                GetInCellValidationSelectedIndex = i
                Exit Function
            End If
        Next
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Graphic/Picture Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function CopyPictureShapeToWorksheet(wksDestination As Excel.Worksheet, strFilePath As String, Left As Long, Top As Long, Width As Long, Height As Long) As Shape
'Add picture with unique name using GUID.
'Saves with worksheet.
'    wksDestination.Pictures.Delete
    Dim objPic As Shape
    Dim strFileName As String

    If Dir(strFilePath) <> vbNullString Then
        strFileName = ParsePath(strFilePath, FileName)
        Set objPic = wksDestination.Shapes.AddPicture(strFilePath, msoFalse, msoTrue, Left, Top, Width, Height)
        With objPic
            .LockAspectRatio = msoFalse
            .name = strFileName ' & StGuidGen()
        End With
    End If

    Set CopyPictureShapeToWorksheet = objPic
End Function

Private Function CopyPictureToWorksheet(wksDestination As Excel.Worksheet, strFilePath As String, Left As Long, Top As Long, Width As Long, Height As Long) As Object
'Add picture with unique name using GUID.
'Picture is linked to source document, so if source is deleted, picture link is broken showing no shape.
'    wksDestination.Pictures.Delete
    Dim objPic As Object
    Dim strFileName As String

    If Dir(strFilePath) <> vbNullString Then

        Call SetClipboard(LoadPicture(strFilePath))

        Set objPic = wksDestination.Pictures.Insert(strFilePath) 'Picture Object
        strFileName = ParsePath(strFilePath, FileName)
        With objPic
            .ShapeRange.LockAspectRatio = msoFalse
            .Left = Left
            .Top = Top
            .Height = Height
            .Width = Width
            .name = strFileName ' & StGuidGen()
        End With
    End If

    Set CopyPictureToWorksheet = objPic
End Function

Private Function FindPictureByName(wksSource As Excel.Worksheet, strFilePath As String) As Picture
    On Error GoTo errsub
    Dim strFileName As String

    strFileName = ParsePath(strFilePath, FileName)
    Set FindPictureByName = wksSource.Pictures(strFileName)

errsub:
End Function

Private Function AddPictureGraphicFromWorksheet(wksSource As Excel.Worksheet, wksDestination As Excel.Worksheet, strFile As String, Left As Long, Top As Long, Width As Long, Height As Long)
'Add picture from Graphics worksheet.
    On Error GoTo errsub

    Dim oClipData As Object 'DataObject
    Dim objPicture As Picture

    Set oClipData = GetClipboard 'Backup clipboard contents.

    Set objPicture = FindPictureByName(wksSource, strFile)
    If Not objPicture Is Nothing Then
        objPicture.CopyPicture

        With wksDestination.Pictures.Paste
            .ShapeRange.LockAspectRatio = msoFalse
            .Left = Left
            .Top = Top
            .Height = Height
            .Width = Width
            .name = "CustomPMGraphic" & StGuidGen()
        End With
    End If

errsub:
    Call SetClipboard(oClipData) 'Restore clipboard contents.
End Function

Private Function FindPictureByLocation(wksDestination As Excel.Worksheet, Left As Long, Top As Long, Width As Long, Height As Long) As Object
    On Error GoTo errsub

    Dim obj As Variant
    Dim lbuffer As Long
    lbuffer = 20

    For Each obj In wksDestination.Pictures
        If TypeName(obj) = "Picture" Then
'            If StrComp(VBA.Left(obj.Name, VBA.Len("CustomPMGraphic")), "CustomPMGraphic") = 0 Then 'Same
'            If (VBA.Left(strName, Len("CustomPMGraphic")) = "CustomPMGraphic") Then  'Same
                If (obj.Left > Left - lbuffer And obj.Left < Left + lbuffer) And (obj.Top > Top - lbuffer And obj.Top < Top + lbuffer) And _
                 (obj.Width > Width - lbuffer And obj.Width < Width + lbuffer) And (obj.Height > Height - lbuffer And obj.Height < Height + lbuffer) Then
                    Set FindPictureByLocation = obj
                    Exit Function
                End If
'            End If
        End If
    Next
errsub:
End Function

Private Function AddLineShape(strName As String, wksSheet As Excel.Worksheet, objBegin As Object, objEnd As Object, Optional strLabel As String) As Excel.Shape
'Adds line from left middle of source object to left middle of destination object.
'Returns arrow object.
    On Error GoTo errsub

    Dim obj As Excel.Shape

    'Add arrow
    With wksSheet.Shapes.AddConnector(msoConnectorStraight, objBegin.Left, objBegin.Top + (objBegin.Height) / 2, objEnd.Left, objEnd.Top + (objEnd.Height / 2))
        .name = strName & "_Arrow"
        .Line.EndArrowheadStyle = msoArrowheadOpen
        .Line.ForeColor.RGB = 5066944 'Maroon 'vbRed 'RGB(255, 0, 0)
        .Line.Weight = 2 '1
'        Selection.ShapeRange.ZOrder msoSendToBack
'        Selection.ShapeRange.ZOrder msoBringToFront
        Set AddLineShape = .Line.Parent
    End With

    If strLabel <> vbNullString Then
        'Add Text Box
        With wksSheet.Shapes.AddLabel(msoTextOrientationHorizontal, objBegin.Left, objBegin.Top, 1, 1)
            .name = strName & "_Label"
            .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
            .TextFrame2.WordWrap = msoFalse
            .TextFrame2.TextRange.Text = strLabel
    '       .Fill.Visible = msoTrue 'Solid back solor
            .Left = objBegin.Left - ((objBegin.Left - objEnd.Left) / 2) - (.Width / 2)
            .Top = objBegin.Top - ((objBegin.Top - objEnd.Top) / 2)
        End With

        'Grouped Arrow/Text Box Object
        Set obj = wksSheet.Shapes.Range(Array(strName & "_Arrow", strName & "_Label")).Group
        obj.name = strName & "_Group"

        Set AddLineShape = obj
    End If
    Exit Function
errsub:
    If Err.Number <> 0 Then Err.Raise Err.Number
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Protect & Unprotect Functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'http://datapigtechnologies.com/blog/index.php/hack-into-a-protected-excel-2007-or-2010-workbook/
'http://deinfotech.blogspot.co.uk/2011/08/unprotect-password-protected-excel-2007.html
'http://excelzoom.com/2009/08/how-to-recover-lost-excel-passwords/
'http://download.cnet.com/Free-Word-Excel-and-Password-Recovery-Wizard/3000-2092_4-10249515.html

Private Function WorksheetPasswordBreaker(wksTarget As Excel.Worksheet) As String
    'Code to break a worksheet password.
    'It works because Excel does not use the password entered directly.
    'Instead it is mathematically transformed (hashed) into a much less secure code, being a string of 12 characters, the first 11 of which have one of only two possible values.
    'The remaining character can have up to 95 possible values, leading to only 2^11 * 95 = 194,560 potential passwords!
    'It doesn't matter what your original password is, one of these 194K strings will unlock your sheet or workbook

    Dim i As Integer, j As Integer, k As Integer, l As Integer, m As Integer, n As Integer
    Dim i1 As Integer, i2 As Integer, i3 As Integer, i4 As Integer, i5 As Integer, i6 As Integer
    Dim strPassword As String

    On Error Resume Next
    If wksTarget.ProtectContents = True Then
        For i = 65 To 66: For j = 65 To 66: For k = 65 To 66
        For l = 65 To 66: For m = 65 To 66: For i1 = 65 To 66
        For i2 = 65 To 66: For i3 = 65 To 66: For i4 = 65 To 66
        For i5 = 65 To 66: For i6 = 65 To 66: For n = 32 To 126
            strPassword = VBA.Chr(i) & VBA.Chr(j) & VBA.Chr(k) & VBA.Chr(l) & VBA.Chr(m) & VBA.Chr(i1) & VBA.Chr(i2) & VBA.Chr(i3) & VBA.Chr(i4) & VBA.Chr(i5) & VBA.Chr(i6) & VBA.Chr(n)
            Call wksTarget.Unprotect(strPassword)
            If wksTarget.ProtectContents = False Then
'                Debug.Print strPassword
                Call MsgBox(strPassword)
                Exit Function
            End If
        Next: Next: Next: Next: Next: Next
        Next: Next: Next: Next: Next: Next
    End If
End Function

Private Function WorkbookPasswordBreaker()
    'http://www.youtube.com/watch?v=uWSg-qkCE7c
    'Steps:
    '1. Close the workbook
    '2. Change extension from xlsx or xlsm to zip
    '3. Extract the contents of the zip file
    '4. Go into the xl folder of the extracted file contents
    '5. Edit the workbook.xml file found under the xl folder to remove the following tags:
    '<filesharing>, <workbookPr> and <workbookProtection> (you may not have all of them)
    'e.g. of filesharing tag: <fileSharing reservationPassword="CBEB" userName="Edwin"/>
    'e.g. of workbookPr tag: <workbookPr defaultThemeVersion="124226"/>
    'e.g. of workbookProtection tag: <workbookProtection lockStructure="1" workbookPassword="CBEB"/>
    '6. Overwrite the workbook.xml file in the original zip file with the edited one by doing a simple copy and paste
    '7. Change the file extension from zip back to xlsx or xlsm
    '8. You're good to go!
    '
    'Note:
    '<filesharing> tag is created when you save a readonly
    '<workbookProtection> is created when you protect the structure of the workbook
End Function

Private Function XLSWorkbookPasswordBreakerHEX()
    'A good commercial version available here: http://www.codematic.net/excel-tools/excel-worksheet-password-remover.htm

'    Please try this method for opening password protected workbook.It will surely work i have tested it several times-using a .xls format spreadsheet (the default for Excel up to 2003). For Excel 2007 onwards, the default is .xlsx, which is a fairly secure format, and this method will not work.
'    Backup the xls file
'    Using a HEX editor, locate the DPB=... part
'    Change the DPB=... string to DPx=...
'    Open the xls file in Excel
'    Open the VBA editor (ALT+F11)
'    the magic: Excel discovers an invalid key (DPx) and asks whether you want to continue loading the project (basically ignoring the protection)
'    You will be able to overwrite the password, so change it to something you can remember
'    Save the xls file
'    then open the sheet it is all yours...
End Function

Private Function SetWorksheetProtection(wksTarget As Excel.Worksheet, Protected As Boolean, Optional strPassword As String = "admin") As Boolean
'Sets/removes Worksheet protection
    On Error GoTo errsub

    If Protected = False Then
        If wksTarget.ProtectContents = True Then
            Call wksTarget.Unprotect(strPassword)
        End If
    Else
        Call wksTarget.Protect(strPassword, True, True, True)
    End If
    SetWorksheetProtection = True

errsub:
    If Err.Number = 1004 Then
        Call MsgBox(Err.Description, vbCritical, "Password Error")
    End If
End Function

Private Function ApplyWorkbookContentProtection(rng As Excel.Range, Optional strPassword As String = "admin") As Boolean
'Create protected workbook content
    Dim wks As Excel.Worksheet
    Dim AERange As AllowEditRange
    Dim rngArea As Excel.Range
    Dim i As Long

    Set wks = rng.Parent

    If SetWorksheetProtection(wks, False, strPassword) = True Then
        'Remove previous protection
        For Each AERange In wks.Protection.AllowEditRanges
            AERange.Delete
        Next

        'Add unprotected/editable ranges
        For Each rngArea In rng.Areas
            Call wks.Protection.AllowEditRanges.Add("EditRange_" & i, rngArea)
            i = i + 1
        Next

        'Re-Protect Worksheet
        Call SetWorksheetProtection(wks, True, strPassword)

        ApplyWorkbookContentProtection = True
    End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' Supporting functions
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function ParsePath(ByVal strPath As String, iMode As PathParseMode) As String
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
            If InStr(1, strPath, "\") > 0 Then  '"\" exists
                strPath = fso.GetParentFolderName(strPath) 'File path
                ParsePath = fso.BuildPath(strPath, "\") 'Add "\"
            End If
        Case PathParseMode.FileNameWithoutExtension
            ParsePath = fso.GetBaseName(strPath) 'File name without extension
    End Select
errsub:
    Set fso = Nothing
End Function

Private Sub RemoveLibraryReferences()
'Programmatically remove references from the reference list.
'http://support.microsoft.com/kb/213524
    On Error Resume Next
    Dim xObject As Object
    Set xObject = ThisWorkbook.VBProject.References.Item("Office")
    ThisWorkbook.VBProject.References.Remove xObject
    Set xObject = ThisWorkbook.VBProject.References.Item("stdole")
    ThisWorkbook.VBProject.References.Remove xObject
End Sub

Private Sub EraseRange(rng As Excel.Range)
    On Error Resume Next
    rng.ClearContents
    On Error GoTo 0
End Sub

Private Sub SetFormatToGeneral(wks As Excel.Worksheet)
    'Implemented to work around bug that changes format of a column in the worksheet in focus if worksheet in focus is not the worksheet being manipulated.
    On Error Resume Next
    wks.UsedRange.EntireColumn.NumberFormat = "General"
    On Error GoTo 0
End Sub

Private Function FormatAllExcelStringsForProModelSimple(ByRef rngSource As Excel.Range)
'Faster method - Based on "FormatAllStringsForProModel" which is SLOW!
'http://www.mydigitallife.info/2009/12/27/how-to-find-or-replace-question-mark-asterisk-or-tilde-in-microsoft-office-excel/
'Use "~" in front of "*" or "?" to indicate literal character and not pattern matching.
    On Error GoTo errsub

    Dim i As Long
    Dim vArray As Variant
    Dim rng As Excel.Range
    Dim bEventsPreviousSetting As Boolean
'    Dim bPrevious As Boolean

'    bPrevious = Application.DisplayAlerts

    vArray = Array("+", "-", "~*'", "/", ",", ":", ";", "(", ")", "[", "]", "{", "}", """, """, "<", ">", "=", "\", "\", "'", "!", "@", "#", "$", "%", "^", "&", "|", "~?", "`", "~", ".", vbNullString, vbLf, vbCr, vbTab, VBA.Space(1))

    bEventsPreviousSetting = Application.EnableEvents
    Application.EnableEvents = False
'    Application.DisplayAlerts = False

    For i = 0 To UBound(vArray)
        Call rngSource.Replace(vArray(i), "_", xlPart, Excel.XlSearchOrder.xlByRows, False, , False, False)
'        strData = Replace(strData, CStr(vArray(i)), "_")
    Next

'    'First character can't be number
'    If VBA.Left(strData, 1) Like "[0-9]" Then
'        strData = "_" & VBA.Mid(strData, 2)
'    End If

'    'String length limit
'    If Len(strData) > 74 Then
'        strData = VBA.Left(strData, 74)
'    End If

'    For Each rng In rngSource
'        rng.Value = FormatStringForProModel(CStr(rng.Value))
'    Next

errsub:
'    Application.DisplayAlerts = bPrevious
    If bEventsPreviousSetting <> Application.EnableEvents Then
        Application.EnableEvents = bEventsPreviousSetting
    End If
End Function

Private Function ReIndex(StartNumber As Long, rngDestination As Excel.Range)
    'Pass in range.  Fills with linear trend of numbers starting with StartNumber
    Dim rngHeader As Excel.Range

    rngDestination(1).Value = StartNumber
    Set rngHeader = rngDestination(1).Resize(, rngDestination.Columns.Count)
    Call rngHeader.AutoFill(rngDestination, xlLinearTrend)
End Function

Private Function CopyEntireColumn(rngSource As Excel.Range, rngDestination As Excel.Range) As Excel.Range

    Call GetUsedColumnByStartCell(rngDestination(1)).ClearContents ' rngDestination

    Set rngSource = GetUsedColumnByStartCell(rngSource(1))
    Set rngDestination = rngDestination(1).Resize(rngSource.Rows.Count, rngSource.Columns.Count)
    rngDestination = rngSource.Value
    Set CopyEntireColumn = rngDestination
End Function

Private Sub EatKey(Key As String)
'Note on how to disable a keyboard key.  Example is delete key.
'    Excel.Application.OnKey "{Delete}", vbNullString
'    Excel.Application.OnKey "{Delete}", "ThisWorkbook.EatKey"  'Disable Delete Key
End Sub

Private Sub ToggleRibbon(bVisible As Boolean)
    'Show hide Office 2007 & 2010 Ribbon
    Application.ExecuteExcel4Macro "Show.Toolbar(""Ribbon"", " & bVisible & ")"
End Sub

Private Function GetNumPrintedPages() As Long
    'Get the number of printed pages.
    GetNumPrintedPages = Application.ExecuteExcel4Macro("Get.Document(50)")
End Function

Private Function GetClipboard() As Object 'DataObject 'Requires reference to userforms
    On Error GoTo errsub

    Dim oData As Object 'DataObject 'Requires reference to userforms

    Set oData = GetDataObject()
    Call oData.GetFromClipboard
    Set GetClipboard = oData
errsub:
End Function

Private Function StGuidGen() As String
    'http://www.mrexcel.com/tip078.shtml
    'http://www.cpearson.com/Excel/CreateGUID.aspx
    'Generates a new GUID, returning it in canonical (string) format

    Dim rclsid As GUID

    If CoCreateGuid(rclsid) = 0 Then
        StGuidGen = StGuidFromGuid(rclsid)
    End If
End Function

Private Function StGuidFromGuid(rclsid As GUID) As String
    'http://www.mrexcel.com/tip078.shtml
    'http://www.cpearson.com/Excel/CreateGUID.aspx
    'Converts a binary GUID to a canonical (string) GUID.

    Dim rc As Long
    Dim stGuid As String

    ' 39 chars  for the GUID plus room for the Null char
    stGuid = String$(40, vbNullChar)
    rc = StringFromGUID2(rclsid, StrPtr(stGuid), VBA.Len(stGuid) - 1)
    StGuidFromGuid = VBA.Left(stGuid, rc - 1)
End Function

Private Function SetClipboard(oData As Object) As Boolean
'oData is of type DataObject
    On Error GoTo errsub

    oData.PutInClipboard
    SetClipboard = True

errsub:
End Function

Private Sub GiveExcelFocus()
    'Used to give back focus to Excel when lost.
    Call VBA.Interaction.AppActivate(Application.Caption)
End Sub

Private Function SaveTempExcelWorksheetAsAsCSV(wksSource As Object) As String 'Object = Excel.Worksheet
    On Error GoTo errsub
    Dim strPath As String

    strPath = GetTempFileName("csv")

'    Application.DisplayAlerts = False
    Call wksSource.SaveAs(strPath, xlCSVMSDOS, False)
'    Application.DisplayAlerts = True

    SaveTempExcelWorksheetAsAsCSV = strPath
'    Debug.Print "Saved temporary copy of file connection to: " & strFile
errsub:
End Function

Private Function GetTempFileName(Optional strFileExtension As String = "tmp") As String
'When called using UNIQUE_NAME creates unique temp file name.
'WinAPI GetTempPath() can also be used to get temp path instead of VBA.Environ("TEMP").
'Returns full path name to unique file.
    On Error GoTo errsub
    Dim strResult As String

    strResult = VBA.Space(MAX_PATH)
    Call GetTempFileNameA(VBA.Environ("TEMP"), "TMP", UNIQUE_NAME, strResult)
    strResult = Left$(strResult, InStr(strResult, VBA.Chr(0)) - 1)
    If Dir(strResult) <> vbNullString Then 'File creation ensures unique file name.
        Call Kill(strResult)
        strFileExtension = VBA.Replace(strFileExtension, ".", "")
        GetTempFileName = VBA.Left(strResult, Len(strResult) - 3) & strFileExtension
    End If
errsub:
End Function

Private Sub DumpDataFromInternetExplorer()
'Quick and dirtly data dump from web site to Excel.
'Untested
    Dim ie As Object
    Dim RowCount As Long
    Dim URL As String
    Dim itm As Variant

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True

    URL = "http://www.sgx.com/wps/portal/sgxweb/home/company_disclosure/company_announcements"

    'Wait for site to fully load
    ie.Navigate2 URL
    Do While ie.Busy = True
        DoEvents
    Loop

    RowCount = 1

    With Sheets("Sheet1")
        .Cells.ClearContents
        RowCount = 1
        For Each itm In ie.Document.all
            .Range("A" & RowCount) = Left(itm.innertext, 1024)
            RowCount = RowCount + 1
        Next
    End With
End Sub

Private Sub ImportPDFFileData(strPDFFileName As String)
    'Untested
    'How To Import PDF File Data Into Excel Worksheet
    'Author: Steve Lipsman
    'Purpose: Import PDF File Data Into Excel Worksheet
    'Other Requirement(s): 'Acrobat' Checked in VBA Tools-References
    'Reference Renames Itself 'Adobe Acrobat 9.0 Object Library' After Reference Is Saved

    On Error GoTo errsub

    'Declare Variable(s)
    Dim appAA As Object 'Acrobat.CAcroApp
    Dim docPDF As Object 'Acrobat.CAcroPDDoc
    Dim strFileName As String
    Dim intNOP As Long
    Dim arrI As Variant
    Dim intC As Long
    Dim intR As Long
    Dim intBeg As Long
    Dim intEnd As Long

    'Initialize Variables
    Set appAA = CreateObject("AcroExch.App")
    Set docPDF = CreateObject("AcroExch.PDDoc")

    'Read PDF File
    Call docPDF.Open(strFileName)

    'Extract Number of Pages From PDF File
    intNOP = docPDF.GetNumPages

    'Select First Data Cell
    Range("A1").Select

    'Open PDF File
    ActiveWorkbook.FollowHyperlink strFileName, , True

    'Loop Through All PDF File Pages
    For intC = 1 To intNOP
        'Go To Page Number
        SendKeys ("+^n" & intC & "{ENTER}")

        'Select All Data In The PDF File's Active Page
        SendKeys ("^a"), True

        'Right-Click Mouse
        SendKeys ("+{F10}"), True

        'Copy Data As Table
        SendKeys ("c"), True

        'Minimize Adobe Window
        SendKeys ("%n"), True

        'Paste Data In This Workbook's Worksheet
        ActiveSheet.Paste

        'Select Next Paste Cell
        Range("A" & Range("A1").SpecialCells(xlLastCell).Row + 2).Select

        'Maximize Adobe Window
        SendKeys ("%x")
    Next

    'Close Adobe File and Window
    SendKeys ("^w"), True

errsub:
    'Empty Object Variables
    Set appAA = Nothing
    Set docPDF = Nothing
End Sub

Public Function GetRangeQueryAddress(rngData As Excel.Range) As String
    If Not rngData Is Nothing Then
        GetRangeQueryAddress = "[" & rngData.Parent.name & "$" & rngData.Address(False, False) & "]"
    End If
End Function

Private Function ShowMessageBoxForTime(strMessage As String, Timer As Long)
    Call CreateObject("Wscript.Shell").Popup(strMessage, Timer)
'    Call CreateObject("Wscript.Shell").Popup("Limited Time popup messagebox for 1 second.", 1)
End Function

Private Sub ListExcelReferencePaths(oThisWorkbook As Object)
'Determines the full path and Globally Unique Identifier (GUID)
'to each referenced library.  Select the reference in the Tools\References
'window, then run this code to get the information on the reference's library
'Requires trust to VBA Project Object model
    On Error Resume Next
    Dim i As Long
    
    
    For i = 1 To oThisWorkbook.VBProject.References.Count
        With ThisWorkbook.VBProject.References(i)
            Debug.Print "Name:" & .name & " Path:" & .FullPath & " GUID:" & .GUID
        End With
    Next
    On Error GoTo 0
End Sub

Private Function GetDataObject() As Object
    'Create new object by class GUID
    'This is ok when you use a UserForm referenced with "Microsoft Forms 2.0 Object Library" - FM20.dll
    'Dim oDataObject As MSForms.DataObject
    'Set oDataObject = New MSForms.DataObject
    
    Set GetDataObject = CreateObject("new:{1C3B4210-F441-11CE-B9EA-00AA006B1A69}") 'Microsoft.Vbe.Interop.Forms.DataObjectClass
'    Set GetDataObject = CreateObject("Application.Forms")
End Function
