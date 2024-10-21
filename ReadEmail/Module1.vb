Imports Microsoft.Office.Interop.Outlook

Module Module1

    WithEvents dgMail As Outlook.MailItem

    Private sentToMe As Outlook.Folder
    Private vipFolder As Outlook.Folder
    Private sentToMyGroup As Outlook.Folder
    Private verifyingFolder As Outlook.Folder
    Private activeFolder As Outlook.Folder
    Private backlogFolder As Outlook.Folder
    Private newFolder As Outlook.Folder
    Private forFollowUpFolder As Outlook.Folder
    Sub Main()

        Dim olApp As Outlook.Application  'Riferimento processo OUTLOOK
        Dim olNs As Outlook.NameSpace     'Namespace utilizzato per navigare i pst, ottenere la sessione, gli elementi selezionati...

        olApp = New Outlook.Application
        olNs = olApp.GetNamespace("MAPI")
        Dim defaultStore As Outlook.Store = olNs.DefaultStore
        Dim inboxFolder As Outlook.Folder = defaultStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

        Dim searchFolderPrefix As String = defaultStore.GetRootFolder.FolderPath & "\search folders\"
        For Each curFolder As Outlook.Folder In olNs.DefaultStore.GetSearchFolders

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Sent Straight to me") Then
                sentToMe = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "VIP") Then
                vipFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Sent To My Groups") Then
                sentToMyGroup = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Verifying") Then
                verifyingFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Active") Then
                activeFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Backlog") Then
                backlogFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "New") Then
                newFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "For Follow Up") Then
                forFollowUpFolder = curFolder
            End If

        Next

        DisplayEmail()


        Console.ReadLine()

    End Sub

    Sub DisplayEmail()
        Console.WriteLine("displayEmail")
        If sentToMe.Items.Count > 0 Then
            Console.WriteLine("sentToMe" & sentToMe.Items.Count)
            dgMail = sentToMe.Items(1)
        ElseIf vipFolder.Items.Count > 0 Then
            Console.WriteLine("vipFolder" & vipFolder.Items.Count)
            dgMail = vipFolder.Items(1)
        ElseIf sentToMyGroup.Items.Count > 0 Then
            Console.WriteLine("sentToMyGroup" & sentToMyGroup.Items.Count)
            dgMail = sentToMyGroup.Items(1)
        ElseIf verifyingFolder.Items.Count > 0 Then
            Console.WriteLine("verifyingFolder" & verifyingFolder.Items.Count)
            dgMail = verifyingFolder.Items(1)
        ElseIf activeFolder.Items.Count > 0 Then
            Console.WriteLine("activeFolder" & activeFolder.Items.Count)
            dgMail = activeFolder.Items(1)
        ElseIf backlogFolder.Items.Count > 0 Then
            Console.WriteLine("backlogFolder" & backlogFolder.Items.Count)
            dgMail = backlogFolder.Items(1)
        ElseIf newFolder.Items.Count > 0 Then
            Console.WriteLine("newFolder" & newFolder.Items.Count)
            dgMail = newFolder.Items(1)
        ElseIf forFollowUpFolder.Items.Count > 0 Then
            Console.WriteLine("forFollowUpFolder" & forFollowUpFolder.Items.Count)
            dgMail = forFollowUpFolder.Items(1)
        End If
        If dgMail Is Nothing Then
            Console.WriteLine("Nothing to Display")
        Else
            dgMail.Display()
        End If

    End Sub

    Sub dgMail_Close(ByRef Cancel As Boolean) Handles dgMail.Close
        Console.WriteLine("Mail Close")

        DisplayEmail()
    End Sub



End Module
