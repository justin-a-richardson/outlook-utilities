Public Sub empty_junk_and_trash_folders()

    Dim fldr As Outlook.Folder
    
    'Delete everything in the junk folder first because they will go to trash
    Set fldr = Application.Session.GetDefaultFolder(olFolderJunk)
    delete_all_items_in_folder fldr
    
    'Delete everything in the deleted items folder now
    Set fldr = Application.Session.GetDefaultFolder(olFolderDeletedItems)
    delete_all_items_in_folder fldr
    
    Set fldr = Nothing

End Sub


Private Sub delete_all_items_in_folder(fldr As Outlook.Folder)

    Dim i As Long
    
    'We count backwards because as we delete items the number changes
    For i = fldr.Items.Count To 1 Step -1
        fldr.Items.Item(i).Delete
    Next i

End Sub


Public Sub mark_all_as_read_in_current_folder()

    Dim fldr As Outlook.Folder
    Dim obj As Object
    
    Set fldr = Application.ActiveExplorer.CurrentFolder
    
    For Each obj In fldr.Items
        If TypeName(obj) = "MailItem" Or TypeName(obj) = "MeetingItem" Then
            'We check the unread status of the object first. We could just mark all items as read
            'but the set function for that flag is very slow. By calling the get function first
            'and making our own determination we actually go a lot faster.
            If obj.UnRead = True Then obj.UnRead = False
        End If
    Next obj

End Sub
