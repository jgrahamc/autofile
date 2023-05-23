Attribute VB_Name = "mdlLearn"
Option Explicit
Public Const INFINITE = &HFFFF      '  Infinite timeout
Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const SW_MINIMIZE = 6
Public Const STARTF_USESHOWWINDOW = &H1
Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "kernel32" _
      (ByVal hObject As Long) As Long
Public Sub Learn(frm As Form)
    Dim otl As Outlook.Application
    
    ' First find the Outlook instance
    On Error GoTo FailedToFindOutlook
    Call SetStatus(frm, "Attaching to Microsoft Outlook...")
    Set otl = New Outlook.Application
    On Error GoTo 0
    
    ' Then read the list of folders
    On Error GoTo FailedToGetFolders
    Call SetStatus(frm, "Getting the list of mail folders...")
    Call GetFolderList(frm, otl)
    On Error GoTo 0
    
    ' The extract from each folder the text of each email creating a collection
    ' of folders as we do
    'On Error GoTo FailedToExtractText
    Call SetStatus(frm, "Extracting email text from folders...")
    Call DeleteDirectory(App.path & "\extract")
    Call ExtractEmailText(frm, otl)
    On Error GoTo 0
    
    ' Release the Outlook instance
    Call SetStatus(frm, "Disconnecting from Microsoft Outlook...")
    Set otl = Nothing
    
    ' Then run the Rainbow command on the extract files
    Call SetStatus(frm, "Learning the structure of all emails...")
    Call RainbowTokenize
    
    ' Clean up the extracted files
    Call SetStatus(frm, "Tidying up temporary files...")
    Call DeleteDirectory(App.path & "\extract")
    
    Call SetStatus(frm, "Done")
    Exit Sub
FailedToFindOutlook:
    Call SetError(frm, "Failed to attach to Microsoft Outlook")
    Exit Sub
FailedToGetFolders:
    Call SetError(frm, "Failed to retrieve list of mail folders")
    Exit Sub
FailedToExtractText:
    Call SetError(frm, "Failed while extracting email text")
    Exit Sub
End Sub
Private Sub SetStatusLabel(frm As Form, status As String, color As Long)
    frm.lblStatus.Caption = status
    frm.lblStatus.ForeColor = color
End Sub
Private Sub SetStatus(frm As Form, status As String)
    Call SetStatusLabel(frm, status, vbBlack)
End Sub
Private Sub SetError(frm As Form, status As String)
    Call SetStatusLabel(frm, status, vbRed)
End Sub
Private Sub GetFolderList(frm As Form, otl As Outlook.Application)
    Dim ns As NameSpace
    Dim inbox As MAPIFolder
    
    ' Retrieve the default MAPI namespace
    Set ns = otl.GetNamespace("MAPI")
    
    ' Retrieve the Inbox
    Set inbox = ns.GetDefaultFolder(olFolderInbox)
    
    ' Walk the Inbox and its folders to build the treeview
    Call RecurseFolderTree(frm, inbox, "")
    
    Set inbox = Nothing
    Set ns = Nothing
End Sub
Private Sub RecurseFolderTree(frm As Form, folder As MAPIFolder, parent As String)
    Dim subFolder As MAPIFolder
    Dim folders As folders
    
    Call AddFolder(frm, folder.name, folder.EntryID, parent)
    
    Set folders = folder.folders
    Set subFolder = folders.GetFirst
    
    While (Not subFolder Is Nothing)
        Call RecurseFolderTree(frm, subFolder, subFolder.parent.EntryID)
        Set subFolder = folders.GetNext
    Wend
End Sub
Private Sub AddFolder(frm As Form, name As String, id As String, parent As String)
    Dim node As node
    
    If (parent <> "") Then
        Set node = frm.tvwFolders.Nodes.Add(parent, tvwChild, id, name)
    Else
        Set node = frm.tvwFolders.Nodes.Add(, , id, name)
    End If
    
    node.Expanded = True
End Sub
Private Sub ExtractEmailText(frm As Form, otl As Outlook.Application)
    Call WalkFolderTree(frm, otl)
End Sub
Private Sub WalkFolderTree(frm As Form, otl As Outlook.Application)
    Dim node As node
    Dim i As Integer
    Dim ns As NameSpace
    
    On Error Resume Next
    Call DeleteDirectory(App.path & "\extract")
    Call MkDir(App.path & "\extract")
    On Error GoTo 0
    
    ' Retrieve the default MAPI namespace
    Set ns = otl.GetNamespace("MAPI")
    
    For i = 1 To frm.tvwFolders.Nodes.count
        Dim name As String
        Set node = frm.tvwFolders.Nodes.item(i)
        node.Bold = True
        name = node.Text
        node.Text = name & " (extracting)"
        
        Call ReadEmailText(frm, otl, node, ns)
        
        node.Text = name & " (100% extracted)"
        node.Bold = False
        Set node = Nothing
        
        DoEvents
    Next
    
    Set ns = Nothing
End Sub
Private Sub ReadEmailText(frm As Form, otl As Outlook.Application, node As node, ns As NameSpace)
    Dim folder As MAPIFolder
    Dim item As MailItem
    Dim count As Long
    Dim progress As Long
    Dim handle As Integer
    Dim name As String
    
    name = node.Text
    handle = FreeFile
    
    On Error Resume Next
    Call MkDir(App.path & "\extract\" & node.key)
    On Error GoTo 0
    Open App.path & "\extract\" & node.key & "\email" For Output As #handle
    
    Call node.EnsureVisible
    
    ' Retrieve the folder
    Set folder = ns.GetFolderFromID(node.key)
    count = folder.Items.count
    If (count > 0) Then
        progress = 0
        While (progress < count)
            Set item = folder.Items.item(progress + 1)
            node.Text = name & " (" & Int(progress * 100 / count) & "% extracted)"
            progress = progress + 1
            Print #handle, "From: " & item.SenderName
            Print #handle, "Subject: " & item.Subject
            Print #handle, item.Body
            DoEvents
        Wend
    End If
    
    node.Text = name & " (100% extracted)"
    
    Close #handle
    
    Set folder = Nothing
End Sub
Private Sub RainbowTokenize()
    On Error Resume Next
    Call MkDir(App.path & "\model")
    On Error GoTo 0
    Call SpawnCommand("-d model --skip-html --index extract/*")
End Sub
Private Sub SpawnCommand(command As String)
    Dim NameOfProc As PROCESS_INFORMATION
    Dim NameStart As STARTUPINFO
    Dim path As String
    path = App.path
    command = Chr(34) & App.path & "\rainbow.exe" & Chr(34) & " " & command
    Dim X As Long
    NameStart.cb = Len(NameStart)
    NameStart.dwFlags = STARTF_USESHOWWINDOW
    NameStart.wShowWindow = SW_MINIMIZE
    X = CreateProcessA(vbNullString, command, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, path, NameStart, NameOfProc)
    X = WaitForSingleObject(NameOfProc.hProcess, INFINITE)
    X = CloseHandle(NameOfProc.hProcess)
    X = CloseHandle(NameOfProc.hThread)
End Sub
Private Sub DeleteDirectory(path As String)
    Dim subdir As String
    
    ' Kill all subdirectories (one level)
    subdir = Dir(path & "\*", vbDirectory)
    
    While (subdir <> "")
        If ((subdir <> ".") And (subdir <> "..")) Then
            Call DeleteEmail(path & "\" & subdir)
            On Error Resume Next
            Call RmDir(path & "\" & subdir)
            On Error GoTo 0
        End If
        subdir = Dir()
    Wend
    
    ' Delete the directory
    On Error Resume Next
    Call RmDir(path)
    On Error GoTo 0
End Sub
Private Sub DeleteEmail(path As String)
    On Error Resume Next
    Kill path & "\email"
    On Error GoTo 0
End Sub
Public Sub SaveAutoFileLocation()
    Call SaveSetting(App.ProductName, "Location", "Path", App.path & "\")
End Sub
