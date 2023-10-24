Public Sub GetEmails()
    Dim objOL As Outlook.Application
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.MAPIFolder
    Dim objMsg As Outlook.MailItem 'Object
    Dim objAttachments As Outlook.Attachments
    Dim i As Long
    Dim lngCount As Long
    Dim strFile As String
    Dim strFolderpath As String
    Dim fso As Object
    
    ' Get the path to your My Documents folder
    strFolderpath = CreateObject("WScript.Shell").SpecialFolders(16)
    On Error Resume Next
    
    ' Instantiate a FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Build the path to the Email_Downloads\ folder
    strFolderpath = fso.BuildPath(strFolderpath, "Email_Downloads")
    
    ' If the directory doesn't exist, create it
    If Not fso.FolderExists(strFolderpath) Then
        fso.CreateFolder strFolderpath
    End If
    
    ' Instantiate an Outlook Application object.
    Set objOL = CreateObject("Outlook.Application")
    
    ' Get the MAPI namespace
    Set objNamespace = objOL.GetNamespace("MAPI")
    
    Dim objStore As Outlook.Store
    For Each objStore In objNamespace.Stores
        If objStore.DisplayName = "your_email@your_domain.designation" Then
            ' Adjust the folder path as needed if you want just inbox stop it at inbox if you need a subfolder continue using it like this
            Set objFolder = objStore.GetRootFolder.Folders("Inbox").Folders("Sub_Zero_Folder")
            Exit For
        End If
    Next objStore
    
    If objFolder Is Nothing Then
        MsgBox "Could not find the folder."
        Exit Sub
    End If
    
    ' Check each selected item for attachments.
    For Each objMsg In objFolder.Items
        Set objAttachments = objMsg.Attachments
        lngCount = objAttachments.Count
        
        If lngCount > 0 Then
            ' Use a count down loop for removing items
            ' from a collection. Otherwise, the loop counter gets
            ' confused and only every other item is removed.
            For i = lngCount To 1 Step -1
                ' Get the file name.
                strFile = objAttachments.Item(i).FileName
                
                ' Combine with the path to the WorkedFile folder.
                strFile = fso.BuildPath(strFolderpath, strFile)
                
                ' Save the attachment as a file.
                objAttachments.Item(i).SaveAsFile strFile
            Next i
        End If
    Next
    
ExitSub:
    Set fso = Nothing
    Set objAttachments = Nothing
    Set objMsg = Nothing
    Set objFolder = Nothing
    Set objNamespace = Nothing
    Set objOL = Nothing
End Sub

