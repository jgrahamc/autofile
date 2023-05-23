VERSION 5.00
Begin {AC0714F6-3D04-11D1-AE7D-00A0C90F26F4} Connect 
   ClientHeight    =   9945
   ClientLeft      =   1740
   ClientTop       =   1545
   ClientWidth     =   6585
   _ExtentX        =   11615
   _ExtentY        =   17542
   _Version        =   393216
   Description     =   "AutoFile mail filer"
   DisplayName     =   "AutoFile"
   AppName         =   "Microsoft Outlook"
   AppVer          =   "Microsoft Outlook 10.0"
   LoadName        =   "Startup"
   LoadBehavior    =   3
   RegLocation     =   "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook"
End
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'----------------------------------------------------------------------
' AutoFile - knn mail filer for Outlook
'
' Copyright (c) 2002 John Graham-Cumming
'
'----------------------------------------------------------------------
Dim IsLicensed As Boolean
Dim HaveNagged As Boolean

' We access the main Outlook application and are interested in its ItemSend event below
Dim WithEvents oXL As Outlook.Application
Attribute oXL.VB_VarHelpID = -1
Dim WithEvents AutoFileEnableDisable As CommandBarButton
Attribute AutoFileEnableDisable.VB_VarHelpID = -1
Dim WithEvents AutoFileLearn As CommandBarButton
Attribute AutoFileLearn.VB_VarHelpID = -1
Private Sub AddinInstance_OnConnection(ByVal Application As Object, ByVal ConnectMode As AddInDesignerObjects.ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error Resume Next
    Dim ins As Explorer
    Set oXL = Application
     
    ' Get the license state
    IsLicensed = CheckLicenseKey()
    
    ' Set the HaveNagged flag to False so that we will Nag the user on the next Receive if the
    ' product is not licensed
    HaveNagged = False
    
    ' Now need to add the toolbar that's associated with AutoFile
    Set ins = oXL.ActiveExplorer
    If (Not ins Is Nothing) Then
        Dim cb As CommandBar
        Set cb = ins.CommandBars("Standard")
        If (Not cb Is Nothing) Then
            Set AutoFileEnableDisable = cb.Controls.Add(msoControlButton, , , , True)
            With AutoFileEnableDisable
                .Picture = LoadPicture(App.path & "\autofile.bmp")
                .Mask = LoadPicture(App.path & "\automask.bmp")
                .Caption = "&AutoFile"
                .Style = msoButtonIconAndCaption
                .Visible = True
                .BeginGroup = True
            End With
            Set AutoFileLearn = cb.Controls.Add(msoControlButton, , , , True)
            With AutoFileLearn
                .Caption = "&Learn"
                .ToolTipText = "Teach AutoFile"
                .Style = msoButtonCaption
                .Visible = True
            End With
            
            Call CheckAutoFileEnabled
            
            Set cb = Nothing
        End If
        Set ins = Nothing
    End If
End Sub
Private Function AutoFileEnabled() As Boolean
    Dim key As String
    
    key = GetSetting(App.Title, "State", "Enabled")
    
    If (key = "No") Then
        AutoFileEnabled = False
    Else
        AutoFileEnabled = True
    End If
End Function
Private Sub CheckAutoFileEnabled()
    ' Set the states of the two buttons depending on whether AutoFile is enabled or not
    If (AutoFileEnabled()) Then
        AutoFileEnableDisable.State = msoButtonDown
        AutoFileEnableDisable.ToolTipText = "Disable AutoFile"
    Else
        AutoFileEnableDisable.State = msoButtonUp
        AutoFileEnableDisable.ToolTipText = "Enable AutoFile"
    End If
End Sub
Private Sub SetAutoFileEnabled(enabled As Boolean)
    Dim value As String
    
    If (enabled) Then
        value = "Yes"
    Else
        value = "No"
    End If
    
    Call SaveSetting(App.Title, "State", "Enabled", value)
    
    Call CheckAutoFileEnabled
End Sub
' Called when Outlook is quitting.  We just release the Outlook object
Private Sub AddinInstance_OnDisconnection(ByVal RemoveMode As AddInDesignerObjects.ext_DisconnectMode, custom() As Variant)
   On Error Resume Next
   Set oXL = Nothing
End Sub
Private Sub CheckLicense()
    ' If the product has not been licensed then decide whether to nag the user about the product.  This happens
    ' the first time a message is sent for each time Outlook is run
    If (Not IsLicensed) Then
        If (Not HaveNagged) Then
            Call MsgBox("Please consider registering AutoFile.", vbOKOnly, "Unregistered: AutoFile")
            
            ' Don't do this again until the next time Outlook is started
            HaveNagged = True
        End If
    End If
End Sub
Private Function CheckLicenseKey() As Boolean
    Dim key As String
    
    key = GetSetting(App.Title, "License", "Key")
    CheckLicenseKey = False
    
    If (Len(key) = 12) Then
        Dim odds As Integer
        Dim evens As Integer
        Dim product As Integer
        
        odds = Val(Mid$(key, 1, 1)) + Val(Mid$(key, 3, 1)) + Val(Mid$(key, 5, 1)) + Val(Mid$(key, 7, 1)) + Val(Mid$(key, 9, 1)) + Val(Mid$(key, 11, 1))
        evens = Val(Mid$(key, 2, 1)) + Val(Mid$(key, 4, 1)) + Val(Mid$(key, 6, 1)) + Val(Mid$(key, 8, 1)) + Val(Mid$(key, 10, 1))
        
        product = (odds * 3 + evens) Mod 10
        
        If (product = Val(Mid$(key, 12, 1))) Then
            CheckLicenseKey = True
        End If
    End If
End Function
Private Sub AutoFileEnableDisable_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    Call SetAutoFileEnabled(Not AutoFileEnabled())
End Sub
Private Sub LaunchAutoFileLearn()
    frmMain.Show
End Sub
Private Sub AutoFileLearn_Click(ByVal Ctrl As Office.CommandBarButton, CancelDefault As Boolean)
    Call LaunchAutoFileLearn
End Sub
Private Sub oXL_NewMail()
    If (AutoFileEnabled()) Then
        Dim objNameSpace As NameSpace
        Dim inbox As MAPIFolder
        Dim mail As MailItem
        Dim folder As MAPIFolder
    
        Call CheckLicense
        
        ' Get a reference to the Inbox that will now contain new
        ' unread items
        Set objNameSpace = oXL.GetNamespace("MAPI")
        Set inbox = objNameSpace.GetDefaultFolder(olFolderInbox)
        
        ' Iterate through the inbox looking for the unread messages
        For Each mail In inbox.Items
            If mail.UnRead = True Then
                Dim folderid As String
                ' If a message is unread then classify it and get the id of the folder it
                ' should be stored in
                folderid = GetDestinationFolder(mail)
                If (folderid <> "") Then
                    ' Try to map the ID to the actual folder, this could fail
                    ' if the folder has been deleted since we last classified
                    On Error GoTo CouldntFindFolder
                    Set folder = objNameSpace.GetFolderFromID(folderid)
                    On Error GoTo 0
                    
                    ' If we manage to get the folder then move the email
                    If (Not folder Is Nothing) Then
                        Call mail.Move(folder)
                        Set folder = Nothing
                    End If
CouldntFindFolder:
                End If
            End If
        Next
        
        Set inbox = Nothing
        Set objNameSpace = Nothing
        
        Call CheckLastLearn
    End If
End Sub
Private Sub CheckLastLearn()
    Dim lastdate As String
    
    lastdate = GetSetting(App.Title, "State", "LastLearnDate")
    
    If (lastdate = "") Then
        If (MsgBox("Welcome to AutoFile.  Before AutoFile can work correctly it must learn the structure of your email folder.  Would you like AutoFile to learn now?", vbQuestion Or vbYesNo, "AutoFile") = vbYes) Then
            Call LaunchAutoFileLearn
        Else
            Call SetAutoFileEnabled(False)
        End If
        
        Call SaveSetting(App.Title, "State", "LastLearnDate", Month(Date))
    Else
        If (Month(Date) <> lastdate) Then
            If (MsgBox("You have not updated the AutoFile information about your email folders recently. Updating makes AutoFile work better. Would you like AutoFile to learn now?", vbQuestion Or vbYesNo, "AutoFile") = vbYes) Then
                Call LaunchAutoFileLearn
            End If
            
            Call SaveSetting(App.Title, "State", "LastLearnDate", Month(Date))
        End If
    End If
End Sub
Private Function GetDestinationFolder(mail As MailItem) As String
    Dim handle As Integer
    Dim output As String * 1024
    Dim response As String
    Dim filename As String * 1024
    Dim i As Integer
    
    ' Pass the From, Subject and Body of the email into rainbow for classification
    ' and retrieve the id of the mailbox
    
    On Error GoTo fail
    
    ' Save the email message in a temporary file
    handle = FreeFile
    If (GetMailFilename(filename, 1024) > 0) Then
        Open filename For Output As #handle
        Print #handle, mail.Subject
        Print #handle, mail.Body
        Close #handle
        
        ' Classify it
        Call Go(output)
        
        response = ""
        
        ' See what output we got
        For i = 1 To Len(output)
            If (Asc(Mid$(output, i, 1)) > 32) Then
                response = response & Mid$(output, i, 1)
            Else
                Exit For
            End If
        Next
            
        GetDestinationFolder = response
    Else
        GetDestinationFolder = ""
    End If
    Exit Function
fail:
    GetDestinationFolder = ""
End Function
