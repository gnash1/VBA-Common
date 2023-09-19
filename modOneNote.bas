Attribute VB_Name = "modOneNote"
Option Explicit

'Incomplete implementations
'8/8/2023
'https://learn.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh377182(v=office.14)?redirectedfrom=MSDN

'Notebook->Section->Page

Sub SearchTermsInTheFirstNoteBook()
    On Error GoTo errsub:
    
    Dim oneNote As Object ' OneNote.Application
'    Dim notebook As Object ' OneNote.Notebook
    Dim section As Object ' OneNote.Section
    Dim page As Object ' OneNote.Page
    
    ' Connect to OneNote 2010
    ' OneNote will be started if it's not running.
'    Dim oneNote As oneNote.Application 'OneNote14.Application
    
'    Set oneNote = New Microsoft.Office.Interop.oneNote.Application 'OneNote14.Application

    'Check if OneNote is running and create an instance of the OneNote application
'    On Error Resume Next
'    Set oneNote = GetObject(, "oneNote.Application") 'Get existing object in local instance.
    'Set onenoteApp = GetObject(, "OneNote.Application")

    Set oneNote = GetOneNoteObject()
    
    If oneNote Is Nothing Then GoTo errsub
    
'    If Err.Number > 0 Then
'        Err.Clear
'        On Error GoTo 0
'
'        bResult = False
'    Else
'        bResult = True
'    End If

    ' Make the application visible (optional)
'    oneNote.Windows.CurrentWindow.Active = True
'    oneNote.Windows.CurrentWindow.Active = False
    
'    Dim strNotebooksHierarchy As String
'    Call oneNote.GetHierarchy("0", oneNote.HierarchyScope.hsNotebooks, strNotebooksHierarchy)
    
'    oneNote.Visible = True
'    Call oneNote.Publish("piDOC", "C:\Users\Greg\Documents\Downloads", PublishFormat.pfWord)
'    Set oneNote = Nothing
    
    ' Get all of the Notebook nodes.
    Dim NotebookNodes As MSXML2.IXMLDOMNodeList
    Set NotebookNodes = GetOneNoteNotebookNodes(oneNote)
    If Not NotebookNodes Is Nothing Then
        ' Get the first notebook found.
        Dim notebook As MSXML2.IXMLDOMElement
'        Set node = nodes(0)
        For Each notebook In NotebookNodes
            Dim NotebookSections As MSXML2.IXMLDOMSelection
            Set NotebookSections = GetOneNoteNotebookSections(oneNote, notebook)
            
            Dim NotebookSection As Object
            For Each NotebookSection In NotebookSections
                Debug.Print NotebookSection.Name
            
            ' Get the ID.
            Dim notebookID As String
            notebookID = notebook.Attributes.getNamedItem("ID").Text 'GetID Attribute
            
            Debug.Print notebookID
            
            Dim oSection As Object
'            For Each oSection In oneNote.CurrentNotebook.Sections
                Call GetOneNoteNotebookSections(oneNote, notebook)
                ' Export the section to a PDF file.
    '            oSection.ExportToPDF sPath & "\" & oSection.Name & ".pdf"
                Debug.Print oSection.Name
'            Next oSection
            
            
            ' Ask the user for a string for which to search
            ' with a default search string of "Microsoft".
    '        Dim searchString As String
    '        searchString = InputBox$("Enter a search string.", "Search", "Microsoft")
    
    '        Dim searchResultsAsXml As String
            ' The FindPages method search a OneNote object (in this example, the first
            ' open Notebook). You provide the search string and the results are
            ' provided as an XML document listing the objects where the search
            ' string is found. You can control whether OneNote searches non-indexed data (this
            ' example passes False). You can also choose whether OneNote enables
            ' the User Interface to show the found items (this example passes False).
            ' This example instructs OneNote to return the XML data in the 2010 schema format.
    '        oneNote.FindPages notebookID, searchString, searchResultsAsXml, False, False, xs2010
    
            ' Output the returned XML to the Immediate Window.
            ' If no search items are found, the XML contains the
            ' XML hierarchy data for the searched item.
    '        Debug.Print searchResultsAsXml
            Next
        Next
    Else
        MsgBox "OneNote 2010 XML data failed to load."
    End If

errsub:
    Debug.Print Err.Description
    ' Clean up the objects
    Set page = Nothing
    Set section = Nothing
    Set notebook = Nothing
    Set oneNote = Nothing
'    Err.Clear
End Sub

Private Function GetAttributeValueFromNode(node As MSXML2.IXMLDOMNode, attributeName As String) As String
    If node.Attributes.getNamedItem(attributeName) Is Nothing Then
        GetAttributeValueFromNode = "Not found."
    Else
        GetAttributeValueFromNode = node.Attributes.getNamedItem(attributeName).Text
    End If
End Function

Private Function GetHierarchy(oneNote As oneNote.Application) As MSXML2.IXMLDOMNodeList
'Private Function GetFirstOneNoteNotebookNodes(oneNote As OneNote14.Application) As MSXML2.IXMLDOMNodeList
    ' Get the XML that represents the OneNote notebooks available.
    Dim notebooksXml As String
    ' Fill notebookXml with an XML document providing information
    ' about available OneNote notebooks.
    ' To get all the data, provide an empty string
    ' for the bstrStartNodeID parameter.
    Call oneNote.GetHierarchy("", hsNotebooks, notebooksXml, xsCurrent) 'xs2010)

    ' Use the MSXML Library to parse the XML.
    Dim doc As MSXML2.DOMDocument60 'DOMDocument
    Set doc = New MSXML2.DOMDocument60 'MSXML2.DOMDocument
    
    Debug.Print notebooksXml

    If doc.LoadXML(notebooksXml) Then
        Call doc.SetProperty("SelectionNamespaces", "xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote'")
        Set GetFirstOneNoteNotebookNodes = doc.DocumentElement.SelectNodes("//one:Notebook")
    Else
        Set GetFirstOneNoteNotebookNodes = Nothing
    End If
End Function

Private Function GetOneNoteNotebookNodes(oneNote As oneNote.Application) As MSXML2.IXMLDOMNodeList
'Private Function GetFirstOneNoteNotebookNodes(oneNote As OneNote14.Application) As MSXML2.IXMLDOMNodeList
    ' Get the XML that represents the OneNote notebooks available.
    Dim notebooksXml As String
    ' Fill notebookXml with an XML document providing information
    ' about available OneNote notebooks.
    ' To get all the data, provide an empty string
    ' for the bstrStartNodeID parameter.
    Call oneNote.GetHierarchy("", hsNotebooks, notebooksXml, xsCurrent) 'xs2010)

    ' Use the MSXML Library to parse the XML.
    Dim doc As MSXML2.DOMDocument60 'DOMDocument
    Set doc = New MSXML2.DOMDocument60 'MSXML2.DOMDocument
    
    Debug.Print notebooksXml

    If doc.LoadXML(notebooksXml) Then
'        Call doc.SetProperty("SelectionNamespaces", "xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote'")
'        Set OneNoteNotebookNodes = doc.DocumentElement.SelectNodes("//one:Notebook")
        Set GetOneNoteNotebookNodes = doc.DocumentElement.getElementsByTagName("one:Notebook")
    Else
        Set GetOneNoteNotebookNodes = Nothing
    End If
End Function

Private Function GetOneNoteNotebookSections(oneNote As Object, notebook As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMSelection
'    notebook.GetHierarchy(sectionIndex, hsSections, section)
'    Dim section As Object ' Late-bound OneNote.Section
    
    ' Get the XML that represents the OneNote notebooks available.
'    Dim SectionsXml As String
    ' Fill notebookXml with an XML document providing information
    ' about available OneNote notebooks.
    ' To get all the data, provide an empty string
    ' for the bstrStartNodeID parameter.
    
    ' Get the ID.
    Dim notebookID As String
    Dim sectionsXml As String
    notebookID = notebook.Attributes.getNamedItem("ID").Text 'GetID Attribute
            
    Call oneNote.GetHierarchy(notebookID, hsSections, sectionsXml, xsCurrent) 'xs2010)
    '    Dim strNotebooksHierarchy As String
    '    Call oneNote.GetHierarchy("0", oneNote.HierarchyScope.hsNotebooks, strNotebooksHierarchy)

    ' Use the MSXML Library to parse the XML.
'    Dim doc As MSXML2.IXMLDOMNodeList
'    Set doc = New MSXML2.IXMLDOMNodeList
    
'    Debug.Print notebooksXml

'    If doc.LoadXML(notebooksXml) Then
'        Call notebook.ParentNode.SetProperty("SelectionNamespaces", "xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote'")
'        Set GetOneNoteNotebookSections = notebook.SelectNodes("//one:Section")
        Set GetOneNoteNotebookSections = notebook.getElementsByTagName("one:Section")
'    Else
'        Set GetOneNoteNotebookSections = Nothing
'    End If
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
            If UCase(objProc.Name) = UCase(procName) Then  'Double check
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

Sub AccessOneNote16ObjectModel()
    Dim onenoteApp As oneNote.Application ' OneNote.Application 'Object
    Dim onenoteAppIsRunning As Boolean
    Dim notebook As Object ' OneNote.Notebook
    Dim section As Object ' OneNote.Section
    Dim page As Object ' OneNote.Page
    Dim pageContent As Object ' OneNote.PageContent
    Dim pageXML As String

    ' Check if OneNote is running and create an instance of the OneNote application
    On Error Resume Next
    Set onenoteApp = GetObject(, "ONENOTE.APPLICATION") 'Get existing object in local instance.
    'Set onenoteApp = GetObject(, "OneNote.Application")
    
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Set onenoteApp = CreateObject("ONENOTE.APPLICATION") 'Create new object
        onenoteAppIsRunning = False
    Else
        onenoteAppIsRunning = True
    End If

    ' Make the application visible (optional)
'    onenoteApp.Visible = True

    ' Open an existing notebook or create a new one
    Dim strHierarchy As String
    Call onenoteApp.GetHierarchy("0", oneNote.HierarchyScope.hsNotebooks, strHierarchy)
    'Set notebook = onenoteApp.GetHierarchy("0", OneNote.HierarchyScope.hsNotebooks, strTest)

    If notebook Is Nothing Then
        ' Create a new notebook if it doesn't exist
        Set notebook = onenoteApp.CreateNewNotebook("My New Notebook")
    End If

    ' Create a new section in the notebook
    Set section = notebook.AddSection("My New Section")

    ' Create a new page in the section
    Set page = section.AddPage

    ' Set the title of the page
    page.Title = "My New Page"

    ' Add content to the page (e.g., text)
    Set pageContent = page.AddOutline(1)
    pageContent.PageTitle = "My Page Title"
    pageContent.AppendPageText "<p>This is some sample text on the page.</p>"

    ' Get the XML of the page's content
    pageXML = pageContent.XML

    ' Display the XML in a message box (optional)
    MsgBox pageXML

    ' Close the notebook (optional)
    notebook.Close

    ' Quit OneNote if it was not running before running this macro (optional)
    If Not onenoteAppIsRunning Then
        onenoteApp.Quit
    End If

    ' Clean up the objects
    Set pageContent = Nothing
    Set page = Nothing
    Set section = Nothing
    Set notebook = Nothing
    Set onenoteApp = Nothing
End Sub

'Set the locations for notes sent to OneNote in advance.
'(Open your OneNote, click File tab > Options > Send to OneNote)
Private Sub OutlookMacroSample()
  Dim itm As Object
  Const CtrlId As String = "MoveToOneNote"
  For Each itm In ActiveExplorer.Selection
    With itm.GetInspector
      .Display
      If .CommandBars.GetEnabledMso(CtrlId) Then .CommandBars.ExecuteMso CtrlId
      .Close olDiscard
    End With
  Next
End Sub

Private Sub ExportAllOneNoteContentVB6()
    Dim onenoteApp As Object ' Late-bound OneNote.Application
    Dim notebook As Object ' Late-bound OneNote.Notebook
    Dim section As Object ' Late-bound OneNote.Section
    Dim page As Object ' Late-bound OneNote.Page
    Dim notebookPath As String
    Dim outputPath As String
    Dim pageIndex As Long
    Dim sectionIndex As Long
    Dim notebookIndex As Long

    ' Set the file path for the exported content
    outputPath = "C:\Users\YourUsername\Documents\ExportedOneNoteContent\"

    ' Create a new instance of the OneNote application
    Set onenoteApp = CreateObject("OneNote.Application")

    ' Get the hierarchy of notebooks
    notebookIndex = 1
    Do While True
        notebookPath = ""
        
        Dim strNotebookHierarchy As String
        Call onenoteApp.GetHierarchy("0", oneNote.HierarchyScope.hsNotebooks, strNotebookHierarchy)
        'Call GetHierarchy
        
        If onenoteApp.GetHierarchy(notebookPath, 1, notebook) = True Then
        'If onenoteApp.GetHierarchy(notebookPath, 1, notebook) = True Then
            ' Export notebook information
            Debug.Print "Notebook " & notebookIndex & ": " & notebook.Name
            notebook.Export outputPath & "Notebook_" & notebookIndex & "_" & notebook.Name & ".one"

            ' Get sections in the notebook
            sectionIndex = 1
            Do While True
                If notebook.GetHierarchy(sectionIndex, hsSections, section) = True Then
                    ' Export section information
                    Debug.Print " - Section " & sectionIndex & ": " & section.Name
                    section.Export outputPath & "Notebook_" & notebookIndex & "_Section_" & sectionIndex & "_" & section.Name & ".one"

                    ' Get pages in the section
                    pageIndex = 1
                    Do While True
                        If section.GetPage(pageIndex, hsPages, page) = True Then
                            ' Export page content as HTML
                            Debug.Print "    - Page " & pageIndex & ": " & page.Title
                            page.SaveAsHtml outputPath & "Notebook_" & notebookIndex & "_Section_" & sectionIndex & "_Page_" & pageIndex & "_" & page.Title & ".html"
                        Else
                            Exit Do ' No more pages in the section
                        End If
                        pageIndex = pageIndex + 1
                    Loop

                    Set page = Nothing
                Else
                    Exit Do ' No more sections in the notebook
                End If
                sectionIndex = sectionIndex + 1
            Loop

            Set section = Nothing
        Else
            Exit Do ' No more notebooks
        End If
        notebookIndex = notebookIndex + 1
    Loop

    ' Clean up the objects
    Set page = Nothing
    Set section = Nothing
    Set notebook = Nothing
    Set onenoteApp = Nothing

    MsgBox "Export completed!"
End Sub

Private Function GetOneNoteObject() As Object
    On Error Resume Next
    If IsProcessRunning("ONENOTE.EXE") Then
        Set GetOneNoteObject = CreateObject("oneNote.Application") 'Get existing instance (OneNote method is different than other Office applications)
    Else
        Set GetOneNoteObject = New oneNote.Application 'OneNote14.Application
    End If
End Function

'Private Sub GetSectionTest(notebook As MSXML2.IXMLDOMElement)
''    Dim oneNote As Object
'
'    Dim secDoc As MSXML2.DOMDocument60
'    Set secDoc = New MSXML2.DOMDocument60
'
'    Dim secNodes As MSXML2.IXMLDOMNodeList
'    Set secNodes = notebook.DocumentElement.getElementsByTagName("one:Section")
'
'    ' Get the first section.
'    Dim secNode As MSXML2.IXMLDOMNode
'    Set secNode = secNodes(0)
'
'    Dim sectionName As String
'    'sectionName = secNode.Attributes.getNamedItem("name").Text
'    sectionName = secNode.Attributes.getNamedItem(notebookID)
'
'    Dim sectionID As String
'    sectionID = secNode.Attributes.getNamedItem("ID").Text
'
''    oneNote.DeleteHierarchy (sectionID)
''    oneNote.OpenHierarchy
'End Sub

'Add references to
'Microsoft OneNote 14.0 Object Library
'Microsoft XML, 6.0


Public Sub CreateNote()
  ' Connect to OneNote 2010.
    ' To see the results of the code,
    ' you'll want to ensure the OneNote 2010 user
    ' interface is visible.
    Dim oneNote As oneNote.Application
    
    Set oneNote = New oneNote.Application
    
    ' Get all of the Notebook nodes.
    Dim nodes As MSXML2.IXMLDOMNodeList
    Set nodes = GetFirstOneNoteNotebookNodes(oneNote)
    If Not nodes Is Nothing Then
        ' Get the first OneNote Notebook in the XML document.
        Dim node As MSXML2.IXMLDOMNode
        Set node = nodes(0)
        Dim noteBookName As String
        noteBookName = node.Attributes.getNamedItem("name").Text
        
        ' Get the ID for the Notebook so the code can retrieve
        ' the list of sections.
        Dim notebookID As String
        notebookID = node.Attributes.getNamedItem("ID").Text
        
        ' Load the XML for the Sections for the Notebook requested.
        Dim sectionsXml As String
        oneNote.GetHierarchy notebookID, hsSections, sectionsXml, xs2010
        
        Dim secDoc As MSXML2.DOMDocument60
        Set secDoc = New MSXML2.DOMDocument60
    
        If secDoc.LoadXML(sectionsXml) Then
            ' select the Section nodes
            Dim secNodes As MSXML2.IXMLDOMNodeList
            Set secNodes = secDoc.DocumentElement.SelectNodes("//one:Section")
            
            If Not secNodes Is Nothing Then
                ' Get the first section.
                Dim secNode As MSXML2.IXMLDOMNode
                Set secNode = secNodes(0)
                
                Dim sectionName As String
                sectionName = secNode.Attributes.getNamedItem("name").Text
                Dim sectionID As String
                sectionID = secNode.Attributes.getNamedItem("ID").Text
                
                ' Create a new blank Page in the first Section
                ' using the default format.
                Dim newPageID As String
                oneNote.CreateNewPage sectionID, newPageID, npsDefault
                
                ' Get the contents of the page.
                Dim outXML As String
                oneNote.GetPageContent newPageID, outXML, piAll, xs2010
                
                Dim doc As MSXML2.DOMDocument60
                Set doc = New MSXML2.DOMDocument60
                ' Load Page's XML into a MSXML2.DOMDocument object.
                If doc.LoadXML(outXML) Then
                    ' Get Page Node.
                    Dim pageNode As MSXML2.IXMLDOMNode
                    Set pageNode = doc.SelectSingleNode("//one:Page")

                    ' Find the Title element.
                    Dim titleNode As MSXML2.IXMLDOMNode
                    Set titleNode = doc.SelectSingleNode("//one:Page/one:Title/one:OE/one:T")
                    
                    ' Get the CDataSection where OneNote store's the Title's text.
                    Dim cdataChild As MSXML2.IXMLDOMNode
                    Set cdataChild = titleNode.SelectSingleNode("text()")
                    
                    ' Change the title in the local XML copy.
                    cdataChild.Text = "A Page Created from VBA"
                    ' Write the update to OneNote.
                    oneNote.UpdatePageContent doc.XML
                    
                    Dim newElement As MSXML2.IXMLDOMElement
                    Dim newNode As MSXML2.IXMLDOMNode
                    
                    ' Create Outline node.
                    Set newElement = doc.createElement("one:Outline")
                    Set newNode = pageNode.appendChild(newElement)
                    ' Create OEChildren.
                    Set newElement = doc.createElement("one:OEChildren")
                    Set newNode = newNode.appendChild(newElement)
                    ' Create OE.
                    Set newElement = doc.createElement("one:OE")
                    Set newNode = newNode.appendChild(newElement)
                    ' Create TE.
                    Set newElement = doc.createElement("one:T")
                    Set newNode = newNode.appendChild(newElement)
                    
                    ' Add the text for the Page's content.
                    Dim cd As MSXML2.IXMLDOMCDATASection
                    Set cd = doc.createCDATASection("Text added to a new OneNote page via VBA.")

                    newNode.appendChild cd
                 
                    
                    ' Update OneNote with the new content.
                    oneNote.UpdatePageContent doc.XML
                    
                    ' Print out information about the update.
                    Debug.Print "A new page was created in "
                    Debug.Print "Section " & sectionName & " in"
                    Debug.Print "Notebook " & noteBookName & "."
                    Debug.Print "Contents of new Page:"
                    
                    Debug.Print doc.XML
                End If
            Else
                MsgBox "OneNote 2010 Section nodes not found."
            End If
        Else
            MsgBox "OneNote 2010 Section XML Data failed to load."
        End If
    Else
        MsgBox "OneNote 2010 XML Data failed to load."
    End If
    
End Sub

Private Function GetAttributeValueFromNode(node As MSXML2.IXMLDOMNode, attributeName As String) As String
    If node.Attributes.getNamedItem(attributeName) Is Nothing Then
        GetAttributeValueFromNode = "Not found."
    Else
        GetAttributeValueFromNode = node.Attributes.getNamedItem(attributeName).Text
    End If
End Function

Private Function GetFirstOneNoteNotebookNodes(oneNote As oneNote.Application) As MSXML2.IXMLDOMNodeList
    ' Get the XML that represents the OneNote notebooks available.
    Dim notebookXml As String
    ' OneNote fills notebookXml with an XML document providing information
    ' about what OneNote notebooks are available.
    ' You want all the data and thus are providing an empty string
    ' for the bstrStartNodeID parameter.
    oneNote.GetHierarchy "", hsNotebooks, notebookXml, xs2010
    
    ' Use the MSXML Library to parse the XML.
    Dim doc As MSXML2.DOMDocument60
    Set doc = New MSXML2.DOMDocument60
    
    If doc.LoadXML(notebookXml) Then
        Call doc.SetProperty("SelectionNamespaces", "xmlns:one='http://schemas.microsoft.com/office/onenote/2013/onenote'")
        Set GetFirstOneNoteNotebookNodes = doc.DocumentElement.SelectNodes("//one:Notebook")
    Else
        Set GetFirstOneNoteNotebookNodes = Nothing
    End If
End Function

