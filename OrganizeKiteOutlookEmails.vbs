
' I recently discovered that the Windows version of Outlook supports vb scripts
' in the run rules once some registry keys are added.  To help organize my
' conversations with Kite through their ticketing system, I wrote this
' to move the emails out of Inbox and SentItems and into folders based on
' kite's ticket numbers.

' This function is called from the rule side of Outlook.

Sub OrganizeKiteTickets(Item As Outlook.MailItem)
    Dim objOutlook As Object
    Dim objNamespace As Outlook.NameSpace
    Dim objFolder As Outlook.Folder
    Dim objKiteFolder As Outlook.Folder
    Dim objDestinationFolder As Outlook.Folder

    Dim strTicketNumber As String
    strTicketNumber = FindTicketNumber(Item.Subject)
    
    Const olFolderInbox = 6
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    Set objKiteFolder = objFolder.Folders("Kite Tickets")
    If Not (FindFolder(objKiteFolder.Folders, strTicketNumber)) Then
        Set objDestinationFolder = objKiteFolder.Folders.Add(strTicketNumber, olFolderInbox)
    Else
        Set objDestinationFolder = objKiteFolder.Folders(strTicketNumber)
    End If
    MsgBox (strTicketNumber & Item.Subject & Item.SenderEmailAddress)
    
    Item.Move (objDestinationFolder)
End Sub

Function FindTicketNumber(Item As String)
    Dim iStartPosition As Integer
    Dim iStopPosition As Integer
    Dim iLength As Integer
    Dim strMatch As String
    
    strMatch = "Ticket#"
    iStartPosition = InStr(1, Item, strMatch, 1) + Len(strMatch)
    iStopPosition = InStr(iStartPosition, Item, "/", 1)
    iLength = iStopPosition - iStartPosition
    FindTicketNumber = Mid(Item, iStartPosition, iLength)
End Function

Function FindFolder(oFolders As Outlook.Folders, sName As String)
    Const FOUND = 0

    If Len(Trim(sName)) = 0 Then Exit Function
    
    Dim bFound As Boolean
    bFound = False
     
    For Each oFolder In oFolders
        If MATCH = StrComp(oFolder.Name, sName) Then
        bFound = True
            Exit For
        End If
    Next
End Function

Function IsSentEmailKiteTicketReply(Recipients As Recipients)
    Const MATCH = 0
    Dim bMatch As Boolean
    bMatch = False
    For Each Recipient In Recipients
        If MATCH = StrComp(LCase(Recipient.Address), "support@kitetechgroup.com") Then
            bMatch = True
            Exit For
        End If
    Next
    IsSentEmailKiteTicketReply = bMatch
End Function

Function IsFromKiteSupport(strSenderAddress As String)
    Const MATCH = 0
    Dim bMatch As Boolean
    bMatch = False
    If MATCH = StrComp(LCase(strSenderAddress), "support@kitetechgroup.com") Then
        bMatch = True
    End If
    IsFromKiteSupport = bMatch
End Function

' I was hoping to do the same thing with Sent Items that I could do with Rules in Outlook
' but apparently rules can only apply to incoming emails, not sent ones.

Sub MoveKiteSentItems()
    Const olFolderSentMail = 5
    Const olFolderInbox = 6
    
    Dim objOutlook As Object
    Dim objNamespace As Outlook.NameSpace
    Dim objSentItemsFolder As Outlook.Folder
    Dim objInboxFolder As Outlook.Folder
    Dim objKiteFolder As Outlook.Folder
    Dim objDestinationFolder As Outlook.Folder
    ' We don't declare Item as a Outlook.MailItem because an Inbox can contain
    ' other types of objects.
    Dim Item As Object
    Dim Items As Outlook.Items
    Dim strTicketNumber As String
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objSentItemsFolder = objNamespace.GetDefaultFolder(olFolderSentMail)
    Set objInboxFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    Set objKiteFolder = objInboxFolder.Folders("Kite Tickets")
    Set Items = objSentItemsFolder.Items
    
    For Each Item In Items
        If Outlook.olMail = Item.Class Then
            If InStr(1, Item.Subject, "oject Ticket") > 1 Then
                If IsSentEmailKiteTicketReply(Item.Recipients) Then
                    strTicketNumber = FindTicketNumber(Item.Subject)
                    Item.Move (objKiteFolder.Folders(strTicketNumber))
                End If
            End If
        End If
    Next
    
End Sub

' For some odd reason the rule I created in Outlook tht called OrganizeKiteTickets(Item As Outlook.MailItem)
' wasn't matching all the emails in Inbox, so I've ended up writing this to move the rest of the emails.
' I originally attempted to call the same OrganizeKiteTickets function as the rule did to see, but
' I kept getting "Object doesn't support this property or method 438" errors even though TypeName(Item)
' showed it was indeed a MailItem. :/

Sub MoveKiteInboxItems()
    Const olFolderInbox = 6
    
    Dim objOutlook As Object
    Dim objNamespace As Outlook.NameSpace
    Dim objSentItemsFolder As Outlook.Folder
    Dim objInboxFolder As Outlook.Folder
    Dim objKiteFolder As Outlook.Folder
    Dim objDestinationFolder As Outlook.Folder
    Dim Item As Object
    Dim Items As Outlook.Items
    Dim strTicketNumber As String
    
    Set objOutlook = CreateObject("Outlook.Application")
    Set objNamespace = objOutlook.GetNamespace("MAPI")
    Set objInboxFolder = objNamespace.GetDefaultFolder(olFolderInbox)
    Set objKiteFolder = objInboxFolder.Folders("Kite Tickets")
    Set Items = objInboxFolder.Items
    
    For Each Item In Items
        If Outlook.olMail = Item.Class Then
            If InStr(1, Item.Subject, "oject Ticket") > 1 Then
                If IsFromKiteSupport(Item.SenderEmailAddress) Then
                    strTicketNumber = FindTicketNumber(Item.Subject)
                    Item.Move (objKiteFolder.Folders(strTicketNumber))
                End If
            End If
        End If
    Next
    
End Sub

