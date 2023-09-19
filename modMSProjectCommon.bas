Attribute VB_Name = "modMSProjectCommon"
Option Explicit
'File:   modMSProjectCommon
'Author:      Greg Harward
'Contact:     gharward@gmail.com
'Copyright © 2012 ThepieceMaker
'Date:        3/3/09
'
'Summary:
'Common code for MS Project automation.
'Created to interact with MS Project to help with automation of setting data cells and modification.

'Online References:
'http://www.oraxcel.com/projects/sqlxl/help/techniques/query/msproject.html

'Driver for connection to MS Project via ADO is not provided for Office 2007.  Excel automation used instead.
'Dim oConnection as object
'Set oConnection = CreateObject("ADODB.Connection")
'oConnection.Connectionstring = "Provider=Microsoft.Project.OLEDB.9.0;Project Name=" & strExcelSourceFile & ";"
'oConnection.open

'Leasons learned:
'In order to enter start/finish dates without durations being adjusted, set task type to Fixed Duration (see table below), with calculation mode set to manual.
'Moving of summary task rows into the middle of other tasks can cause ID(row number) corruption.  A workaround for this is to insert a blank cell into the desired destination first then move the row to the position below the inserted cell, then delete the cell.
'tsk.ConstraintDate changes based on last changed start/finish cell, so order of input makes a difference.
'Paste in Finish Dates first then Start Dates with Type set to Fixed Duration to get Constraint Date to be set to Start Date.  (Constraint Date assumes value of last pasted value)
'When exporting to PS:
'   Can use Project Manger with *.* to directly bring in MS Project file without Publishing via Project Simulator.  NOt sure what validation is performed.
'   Constraint dates will be used as Start Dates so make sure that they are set correctly. See above.

'Calendar settings:
'                       Days   |   Week    |   Month
' 24x7                24      |   168        |   30.4375 (30.44)
' 9-5                   8        |  40           |  21.74
'-----------------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------------

'MS Project Recalculation Table - MS Project 2003 Step By Step p.141
'Task type setting is down the left side.
'Change type is across the top
'Table holds what will be recalculated.

'Set Task Type                          -----------Change------------
'                            Duration                Units                       Work
'Fixed Duration      Work                    Units                       Units
'Fixed units             Work                    Duration                Duration
'Fixed Work              Units                   Duration                Duration
'-----------------------------------------------------------------------------------------------

'Notes:
'   The tasks collection is maintained in row display order.
'   task.ID = Row Number
'   task.UniqueID = UID

'Date Formats:
'yyyy Year
'q Quarter
'mo, mons = Month
'y Day of year
'd Day
'w Weekday
'ww Week
'h Hour
'm Minute
's Second

Public Const strCalendar = "24 Hours"

Public Enum eDependencyType
    BasedOnPreviousTask
    BasedOnFirstSubTask
End Enum

Private Sub SetStartFinishDatesWithCustomFieldInfo()
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks

    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll 'Loop through tasks.
'        Debug.Print tsk.Name
        If tsk.Summary = False Then
'            Debug.Print tsk.Date1
            tsk.Start = tsk.Date1

'            Debug.Print tsk.Date2
            tsk.Finish = tsk.Date2
        Else
            tsk.Date1 = tsk.Start
            tsk.Date2 = tsk.Finish
        End If
        tsk.ConstraintDate = tsk.Start
    Next
End Sub

Private Sub SetConstraintTypeAndDateForSummaryTasks()
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks

    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll 'Loop through tasks.
'        Debug.Print tsk.Name
        If tsk.Summary = True Then
            tsk.ConstraintDate = tsk.Start
            tsk.ConstraintType = pjSNET
        End If
    Next
End Sub

Private Sub FindOutOfOrderDates()
'For Tibotec: Two types of task problems found
'CTL-LPO-2              'Set finish Date to Start Date + 1 month.
'CTL-FPI+3,LPO-3    'Delete
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks
    Dim cCalendar As MSProject.Calendar
    
    Set cCalendar = ActiveProject.BaseCalendars(strCalendar) '2
    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll 'Loop through tasks.
'        Debug.Print tsk.Name
'        If tsk.Summary = False Then
'            Debug.Print tsk.Date1
            If tsk.Date1 > tskAll(tsk.ID + 1).Date1 And Not tsk.Summary And Not tskAll(tsk.ID + 1).Summary Then
                OutlineShowAllTasks
                Call EditGoTo(tsk.ID)
                
                Debug.Print tsk.name & " Level: " & tsk.OutlineLevel
                If tsk.name = "CTL-LPO-2" Then
                    tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
                ElseIf tsk.name = "CTL-FPI+3,LPO-3" Then
                    tsk.Delete
                End If
            End If
    Next
End Sub

Private Sub FindProblemStartFinishDatesWithConstraintDate()
'For Tibotec: Used to find dates that got adjusted.
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks
    Dim cCalendar As MSProject.Calendar
    
    Set cCalendar = ActiveProject.BaseCalendars(strCalendar) '2
    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll 'Loop through tasks.
        If tsk.Summary = False Then
'           tsk.Type = pjFixedUnits ' pjFixedUnits
'           tsk.ConstraintType = MSProject.PjConstraint.pjSNET
            If tsk.Date1 <> tsk.Start Or tsk.Date2 <> tsk.Finish Or tsk.Start <> tsk.ConstraintDate Then
                Debug.Print tsk.ID
                Debug.Assert False
            End If
        End If
    Next
End Sub

Private Sub FixProblemStartFinishDates2()
'For Tibotec: Used to find dates that got adjusted.
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks
    Dim cCalendar As MSProject.Calendar
    
    Set cCalendar = ActiveProject.BaseCalendars(strCalendar) '2
    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll 'Loop through tasks.
        If tsk.Summary = False Then
'            tsk.Type = pjFixedUnits ' pjFixedUnits
'        End If
            If tsk.Date1 <> tsk.Start Or tsk.Date2 <> tsk.Finish Then
                tsk.ConstraintType = MSProject.PjConstraint.pjSNET
                OutlineShowAllTasks
                Call EditGoTo(tsk.ID)
    '            Debug.Print tsk.Name & " ID: " & tsk.ID
    '        End If
                tsk.Type = pjFixedDuration ' pjFixedUnits ' pjFixedUnits
    '        If tsk.Date2 <> tsk.Finish Then
    '            OutlineShowAllTasks
    '            Call EditGoTo(tsk.ID)
    '            Debug.Print tsk.Name & " ID: " & tsk.ID
    '        End If
    '
    '        If 1 = 2 Then
                If tsk.name = "CTL-LPO-2" Then
                    Dim tDate As Date
    '                tsk.ConstraintDate = "NA"
                    'Oh the insanity!  This is done to work around autocalculation that is going on in MSProject.
                    tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
                    tsk.Finish = tsk.Date2
                    tsk.ConstraintType = MSProject.PjConstraint.pjSNET
                    tsk.Start = tsk.Date1 '
                    tsk.Finish = tsk.Date2
                    tsk.ConstraintDate = tsk.Start 'This needs to be set last for constraint date to get set
    '                Date = Application.DateAdd(tsk.Date1, "1 mons", cCalendar)
                    tsk.ConstraintType = MSProject.PjConstraint.pjSNET
    '                tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
    '                tsk.Finish = tsk.Date2
                    
                ElseIf tsk.name = "CTL-FPI+3,LPO-3" Then
    '                tsk.Delete
                
                Else
                    tsk.Type = pjFixedUnits
    '                tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
'                    tsk.Finish = tsk.Date2
'                    tsk.Start = tsk.Date1
'                    tsk.ConstraintDate = tsk.Start 'This needs to be set first!
                    
                    tsk.Finish = tsk.Date2
                    tsk.ConstraintType = MSProject.PjConstraint.pjSNET
                    tsk.Start = tsk.Date1 '
                    tsk.Finish = tsk.Date2
                    tsk.ConstraintDate = tsk.Start 'This needs to be set last for constraint date to get set
                    tsk.ConstraintType = MSProject.PjConstraint.pjSNET
                    tsk.Type = pjFixedDuration
                    
                End If
            End If
        End If
    Next
End Sub

Private Sub FixCTL()
'For Tibotec: Used to find dates that got adjusted.
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks
    Dim cCalendar As MSProject.Calendar
    
    Set cCalendar = ActiveProject.BaseCalendars(strCalendar) '2
    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll 'Loop through tasks.
        If tsk.name = "CTL-LPO-2" Then
        
            If tsk.Summary = False Then
                tsk.Type = pjFixedUnits ' pjFixedUnits
            End If
            tsk.ConstraintType = MSProject.PjConstraint.pjSNET
'        If tsk.Date1 <> tsk.Start Or tsk.Date2 <> tsk.Finish Then
        
            OutlineShowAllTasks
'            Call EditGoTo(tsk.ID)
'            Debug.Print tsk.Name & " ID: " & tsk.ID
'        End If
'            tsk.Type = pjFixedUnits ' pjFixedUnits
'        If tsk.Date2 <> tsk.Finish Then
'            OutlineShowAllTasks
'            Call EditGoTo(tsk.ID)
'            Debug.Print tsk.Name & " ID: " & tsk.ID
'        End If
'
'        If 1 = 2 Then
            
                Dim tDate As Date
'                tsk.ConstraintDate = "NA"
                
                tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
                tsk.Finish = tsk.Date2
                tsk.Start = tsk.Date1
                
                tsk.ConstraintDate = tsk.Start 'This needs to be set first!
'                Date = Application.DateAdd(tsk.Date1, "1 mons", cCalendar)
'
'                tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
'                tsk.Finish = tsk.Date2
'            ElseIf tsk.Name = "CTL-FPI+3,LPO-3" Then
''                tsk.Delete
'
'            Else
'                tsk.Type = pjFixedUnits
''                tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
'                tsk.Finish = tsk.Date2
'                tsk.Start = tsk.Date1
'
'                tsk.ConstraintDate = tsk.Start 'This needs to be set first!
'            End If
        End If
    Next
End Sub

Private Sub FixProblemStartFinishDates()
'For Tibotec: Two types of task problems found
'CTL-LPO-2              'Set finish Date to Start Date + 1 month.
'CTL-FPI+3,LPO-3    'Delete
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks
    Dim cCalendar As MSProject.Calendar
    
    Set cCalendar = ActiveProject.BaseCalendars(strCalendar) '2
    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll 'Loop through tasks.
'        Debug.Print tsk.Name
'        If tsk.Summary = False Then
'            Debug.Print tsk.Date1
            If tsk.Date1 > tsk.Date2 Then
                OutlineShowAllTasks
                Call EditGoTo(tsk.ID)
                
                Debug.Print tsk.name & " Level: " & tsk.OutlineLevel
                If tsk.name = "CTL-LPO-2" Then
                    tsk.Date2 = Application.DateAdd(tsk.Date1, "1 mons", cCalendar) 'Add 1 month to start date.
                ElseIf tsk.name = "CTL-FPI+3,LPO-3" Then
                    tsk.Delete
                End If
            End If
    Next
End Sub

Private Sub TestCalculateDependancies()
    Call CalculateTaskLevelPredecessors(BasedOnFirstSubTask)
End Sub

Private Sub CalculateTaskLevelPredecessors(DependancyType As eDependencyType) ', units As eUnits) 'Need to add parameter: units calendar
'WARNING: Uses "24 Hour" calendar.
'Code relies on custom fields: UID (tsk.Number1), Original Start (tsk.Date1)
'Two task level assignements used based on level, Provided that tasks are ordered by date within levels (Use SortBy...)
'Type 1. Set dependancy based on common anchor date.
'   Fixes start dates. Tasks could overlap when variability is added.
'Type 2. Set dependancy based on previous task date + lag.
'   When dependancy is added, task start date will move.

    Dim tsk As MSProject.Task
    Dim tskAnchor As MSProject.Task 'Non-Summary task that is one level below summary task.
    Dim tskAll As MSProject.Tasks
    Dim cCalendar As MSProject.Calendar
    
    Set cCalendar = ActiveProject.BaseCalendars(strCalendar) '2
    Set tskAll = ActiveProject.Tasks
    
    'Clear previous
    For Each tsk In tskAll
        tsk.Predecessors = ""
    Next
    
    Select Case DependancyType
        Case eDependencyType.BasedOnFirstSubTask 'Type 1
            For Each tsk In tskAll 'Loop through tasks.
                Debug.Print tsk.name
                If tsk.Summary = True Then
                    If tskAll(tsk.ID + 1).Summary = False Then 'Look at next task.
                        Set tskAnchor = tskAll(tsk.ID + 1)
                    End If
                Else
                    If tskAnchor.UniqueID <> tsk.UniqueID And tsk.Number1 <> 0 Then 'WARNING custom field used here to determine outline level tasks.
                        'Calculate dependancy
                        'WARNING: This code using custom field values!
                        If Application.DateDifference(tskAnchor.Date1, tsk.Date1, cCalendar) * 0.0006944444 < 0 Then
                            OutlineShowAllTasks
                            Call EditGoTo(tsk.ID)
                            MsgBox "Negative lag detected with task: " & tsk.ID
'                            Debug.Assert False
                        Else
'                            tsk.Predecessors = CStr(tskAnchor.ID & "SS+" & DateDiff("d", tskAnchor.Date1, tsk.Date1) & " days") 'Doesn't consider calendars
                            tsk.Predecessors = CStr(tskAnchor.ID & "SS+" & Application.DateDifference(tskAnchor.Date1, tsk.Date1, cCalendar) * 0.0006944444 & " days") 'Considers calendars.  Converts minutes to days.
                        End If
                    End If
                End If
            Next tsk
        Case Else
    End Select
End Sub

Private Sub RemoveAllPredecessors()
    Dim tsk As MSProject.Task
    Dim tskAll As MSProject.Tasks
    
    Set tskAll = ActiveProject.Tasks
    
    For Each tsk In tskAll
        tsk.Predecessors = ""
    Next
End Sub
            
Private Sub CalculatePredecessors()
'First (and subsequent) study task is SS with first (and subsequent) project tasks. Not implemented fully.
'Sets first level 7 Trial date to SS+ lag dependancy on previous first level 5 task.
'Sets subsequent level 7 first in set, trial dates to previous first level 7 + lag.

    Dim tsk As MSProject.Task
    Dim tskLastPrevious As MSProject.Task 'Non-Summary task that is one level below summary task.
    Dim tskAll As MSProject.Tasks
    Dim cCalendar As MSProject.Calendar
    
'    Dim tsk3SummaryAnchor As MSProject.Task 'Task
    Dim tsk3Anchor As MSProject.Task 'Task
    Dim tsk4Anchor As MSProject.Task 'Task
'    Dim tsk4SummaryAnchor As MSProject.Task 'Task
    Dim tsk5Anchor As MSProject.Task 'Task
    Dim tsk5SummaryAnchor As MSProject.Task 'Task
    Dim tsk6Anchor As MSProject.Task 'Task
    Dim tsk6SummaryAnchor As MSProject.Task 'Summary
    Dim tsk7Anchor As MSProject.Task 'Task
    
    Set cCalendar = ActiveProject.BaseCalendars("24 Hours") '2
    
    Set tskAll = ActiveProject.Tasks

        For Each tsk In tskAll 'Loop through tasks.
            If tsk.ID = 495 Then
                Debug.Assert False
            End If
            
            Debug.Print tsk.name
            OutlineShowAllTasks
            Call EditGoTo(tsk.ID) 'Will error if not showing.
            
            'Clear previous value
            tsk.Predecessors = ""
            
'            If tsk.Summary = True Then
                Select Case tsk.OutlineLevel
                    Case 2  'Project

                    Case 3  'Phase
                        If tsk.Summary = False And tskAll(tsk.ID - 1).Summary = True Then 'First one
                            Set tsk3Anchor = tskAll(tsk.ID)
                        ElseIf tsk.Summary = False And tskAll(tsk.ID - 1).Summary = False And tsk.name <> "Phase I" And tsk.name <> "Phase IIA" And tsk.name <> "Phase IIB" And tsk.name <> "Phase III" And tsk.name <> "Launch" Then ' And tskAll(tsk.ID + 1).Summary = False Then 'Look at next task. 2 going to 3
                            tsk.Predecessors = CStr(tsk3Anchor.ID & "SS+" & Application.DateDifference(tsk3Anchor.Date1, tsk.Start, cCalendar) * 0.0006944444 & " days")
                        ElseIf tsk.Summary = True And tskAll(tsk.ID - 1).Summary = False And tskAll(tsk.ID + 1).Summary = False Then 'Study outside phase
                            tsk.Predecessors = CStr(tsk3Anchor.ID & "SS+" & Application.DateDifference(tsk3Anchor.Date1, tsk.Start, cCalendar) * 0.0006944444 & " days")
                        End If
'                        Set tsk4SummaryAnchor = Nothing
                    Case 4  'Study
                        If tsk.Summary = False And tskAll(tsk.ID - 1).Summary = True Then '1st One
                            Set tsk4Anchor = tskAll(tsk.ID)
                        ElseIf tsk.Summary = False And tskAll(tsk.ID - 1).Summary = False Then
                            tsk.Predecessors = CStr(tsk4Anchor.ID & "SS+" & Application.DateDifference(tsk4Anchor.Date1, tsk.Start, cCalendar) * 0.0006944444 & " days")
                        ElseIf tsk.Summary = True And tskAll(tsk.ID + 1).Summary = False Then 'Study within Phase.
                            tsk.Predecessors = CStr(tsk3Anchor.ID & "SS+" & Application.DateDifference(tsk3Anchor.Date1, tsk.Start, cCalendar) * 0.0006944444 & " days")
                        End If
                        Set tsk5SummaryAnchor = Nothing
                    Case 5  'Sub Block
                        '5-6-7
                        
'                        If tsk.Summary = True And tskAll(tsk.ID + 1).Summary = False Then 'And tskAll(tsk.ID + 2).Summary = False Then 'And tsk7Anchor Is Nothing Then  'Look at next task. 6 going to 7
'                            Set tsk6Anchor = tskAll(tsk.ID + 1)
                        If tsk.Summary = False And tskAll(tsk.ID - 1).Summary = True Then 'tsk.UniqueID <> tsk5Anchor.UniqueID Then
                            'First one
                            Set tsk5Anchor = tskAll(tsk.ID)
                        ElseIf tsk.Summary = False And tskAll(tsk.ID - 1).Summary = False Then 'And tskAll(tsk.ID + 1).Summary = False Then 'Last one doesn't get filled if this is set.
                            'Second one in
                            tsk.Predecessors = CStr(tsk5Anchor.ID & "SS+" & Application.DateDifference(tsk5Anchor.Date1, tsk.Start, cCalendar) * 0.0006944444 & " days")
                        ElseIf tsk.Summary = True And tskAll(tsk.ID - 1).Summary = False And tskAll(tsk.ID + 1).Summary = True Then 'Block
                            If Not tsk5SummaryAnchor Is Nothing Then 'skip first block.
                                tsk.Predecessors = CStr(tsk5SummaryAnchor.ID & "FS")
                            End If
                                Set tsk5SummaryAnchor = tsk
                        ElseIf tsk.Summary = True And tskAll(tsk.ID - 1).Summary = True Then '1st Within Block
                            Set tsk5SummaryAnchor = tsk
                        ElseIf tsk.Summary = True And tskAll(tsk.ID - 1).Summary = False Then 'Within Block
                            tsk.Predecessors = CStr(tsk5SummaryAnchor.ID & "FS")
                            Set tsk5SummaryAnchor = tsk
                        End If
                        Set tsk6SummaryAnchor = Nothing
                    Case 6  'Sub Block
                        '6-7
                        If tsk.Summary = False And tskAll(tsk.ID - 1).Summary = True Then 'first
                            Set tsk6Anchor = tskAll(tsk.ID)
                        ElseIf tsk.Summary = False Then
                            tsk.Predecessors = CStr(tsk6Anchor.ID & "SS+" & Application.DateDifference(tsk6Anchor.Date1, tsk.Start, cCalendar) * 0.0006944444 & " days")
                        ElseIf tsk.Summary = True And tskAll(tsk.ID - 1).Summary = True Then '1st Within Block
                            Set tsk6SummaryAnchor = tsk
                        ElseIf tsk.Summary = True And tskAll(tsk.ID - 1).Summary = False Then 'Within Block
                            tsk.Predecessors = CStr(tsk6SummaryAnchor.ID & "FS")
                            Set tsk6SummaryAnchor = tsk
                        End If

                    Case 7
                        If tsk.Summary = False And tskAll(tsk.ID - 1).Summary = True Then 'first
                            Set tsk7Anchor = tskAll(tsk.ID)
                        ElseIf tsk.Summary = False Then
                            tsk.Predecessors = CStr(tsk7Anchor.ID & "SS+" & Application.DateDifference(tsk7Anchor.Date1, tsk.Start, cCalendar) * 0.0006944444 & " days")
                        End If
                    Case Else
                End Select
        Next tsk
End Sub

Private Sub SortByStartDateKeepWithinHeirarchy()
'Same as Project|Sort|SortBy... Sort By StartID, with both check boxes checked.
    Call Sort("Start", , , , , , True, True)
End Sub
    
Private Sub FillSummaryTaskOriginalStartDateInCustomField_Date1()
'Can do 2 levels deep only
    Dim tsk As MSProject.Task
    Dim tskOutlineLevel As MSProject.Task
    Dim tskOutlineLevel2 As MSProject.Task
    Dim fGet As Boolean
    
'    On Error Resume Next ' will error when trying to set predecessor to summary level task.
    
    For Each tsk In ActiveProject.Tasks 'Loop through tasks.
        If tsk.Summary Then
            If fGet = True Then
                Set tskOutlineLevel2 = tsk
            Else
                fGet = True
                Set tskOutlineLevel = tsk
            End If
        Else
            If fGet = True Then
                If Not tsk.Summary Then
                    tskOutlineLevel.Date1 = tsk.Date1
                    If Not tskOutlineLevel2 Is Nothing Then
                        tskOutlineLevel2.Date1 = tsk.Date1
                        Set tskOutlineLevel2 = Nothing
                    End If
                    fGet = False
                End If
            End If
        End If
    Next tsk
    
'    On Error GoTo 0
End Sub

Private Sub FillInPredecessors()
Attribute FillInPredecessors.VB_Description = "Macro Macro2\nMacro Recorded Fri 2/29/08 1:29 AM by Harward."
    Dim tsk As MSProject.Task
'    Set sel = ActiveSelection
'    Set tsk = ActiveSelection.Tasks(1)
    
    On Error Resume Next ' will error when trying to set predecessor to summary level task.
    
    For Each tsk In ActiveProject.Tasks 'Loop through tasks.
        If Not tsk.Summary Then
            If tsk.OutlineLevel > 5 Then
                tsk.Predecessors = tsk.ID - 1
            Else
                tsk.Predecessors = tsk.ID - 1 & "SS"
            End If
        End If
    Next tsk
    
    On Error GoTo 0
End Sub

Private Sub SetConstraintType()
    Dim tsk As MSProject.Task
    
    Set tsk = ActiveSelection.Tasks(1)
    
    On Error Resume Next ' will error when trying to set predecessor to summary level task.
    
    For Each tsk In ActiveProject.Tasks 'Loop through tasks.
        tsk.ConstraintType = MSProject.PjConstraint.pjSNET 'ASAP
    Next tsk
    
    On Error GoTo 0
End Sub

'Based on reply to post on MS Project Developer forum
'http://msdn.microsoft.com/newsgroups/default.aspx?query=elapsed+time&dg=microsoft.public.project.developer&cat=en-us-msdn-officedev-project&lang=en&cr=US&pt=a1d023a3-f612-4da2-acb8-fda8f850d645&catlist=B7714BAA-0D60-40B0-A226-8B9CF33299A5%2C774F24A2-F71F-425F-AC2B-DC48AB0DA5C9&dglist=&ptlist=&exp=&sloc=en-us
Private Sub duration_detector()
    Dim TskTest As MSProject.Task
    Dim iCount As Long
    
    For Each TskTest In ActiveProject.Tasks
        If Not TskTest Is Nothing Then
            If Application.DateDifference(TskTest.Start, TskTest.Finish) = TskTest.Duration Then 'Duration is in working time
                MsgBox TskTest.name
            Else 'Duration is in elapsed time
                iCount = iCount + 1
            End If
        End If
    Next
End Sub

Private Sub Test()
    Dim Test
    Dim TskTest As MSProject.Task
    Dim tskPred As MSProject.Task
    Dim tskTables As MSProject.Tables
    Dim gtskTables
    Dim tdeps As TaskDependencies
    Dim tdep As TaskDependency

    For Each TskTest In ActiveProject.Tasks
        If TskTest.name = "CAN Approval to Phase 1 Start" Then
            Set tdeps = TskTest.TaskDependencies

            For Each tdep In tdeps
'                If tdep.To = tskTest.ID And tdep.From <> earliest.ID Then
                tdep.Lag = 60 * 24 * 10 'minutes 'Allows for setting of Predecessors time, but not sure how to change units.
'                End If
            Next tdep

            Set tskTables = ActiveProject.TaskTables
'            GlobalTaskTables
                Set gtskTables = Application.GlobalTaskTables(Application.ActiveProject.CurrentTable) '"Usage" - 12
'                MsgBox gtskTables.TableFields(5).Value 'pjTaskDuration).Value
'                Set test = Application.ActiveProject.CurrentTable
                
                MsgBox tskTables(7).name
        
'            Beep
            'To see Lag
            Debug.Print TskTest.Predecessors
            For Each tskPred In TskTest.PredecessorTasks
''                Debug.Print
            Next
            'Properties where I can see the lag value: 'Successors, UniqueIDSuccessors, WBSSuccessors
            'Duration
            'Predecessors
            
        End If
    Next
End Sub
