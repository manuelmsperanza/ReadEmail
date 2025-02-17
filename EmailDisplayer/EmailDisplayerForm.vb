Imports Microsoft.Office.Interop.Outlook

Public Class EmailDisplayerForm

    Private olApp As Outlook.Application  'Riferimento processo OUTLOOK
    Private olNs As Outlook.NameSpace     'Namespace utilizzato per navigare i pst, ottenere la sessione, gli elementi selezionati...

    WithEvents dgMail As Outlook.MailItem

    Private sentOnlyToMe As Outlook.Folder
    Private sentToMe As Outlook.Folder
    Private vipFolder As Outlook.Folder
    Private sentToMyGroup As Outlook.Folder
    Private verifyingFolder As Outlook.Folder
    Private activeFolder As Outlook.Folder
    Private backlogFolder As Outlook.Folder
    Private newFolder As Outlook.Folder
    Private forFollowUpFolder As Outlook.Folder

    Private threads As List(Of Thread)
    Private threadIdx As Integer = 0
    Private mailIdx As Integer = 0
    Private Sub EmailDisplayerForm_Shown(sender As Object, e As EventArgs) Handles Me.Shown

        Me.olApp = New Outlook.Application
        Me.olNs = Me.olApp.GetNamespace("MAPI")
        Dim defaultStore As Outlook.Store = Me.olNs.DefaultStore


        Dim searchFolderPrefix As String = defaultStore.GetRootFolder.FolderPath & "\search folders\"
        For Each curFolder As Outlook.Folder In Me.olNs.DefaultStore.GetSearchFolders

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Sent Straight to me") Then
                Me.sentOnlyToMe = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Sent to me") Then
                Me.sentToMe = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "VIP") Then
                Me.vipFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Sent To My Groups") Then
                Me.sentToMyGroup = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Verifying") Then
                Me.verifyingFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Active") Then
                Me.activeFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "Backlog") Then
                Me.backlogFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "New") Then
                Me.newFolder = curFolder
            End If

            If curFolder.FolderPath.Equals(searchFolderPrefix & "For Follow Up") Then
                Me.forFollowUpFolder = curFolder
            End If

        Next

        ' Set AutoSizeColumnsMode for LogDataGridView
        Me.LogDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells

    End Sub

    Sub OrganizeEmail()

        Me.threadIdx = 0
        Me.mailIdx = 0

        Dim threadMap As New Dictionary(Of String, Thread)()

        ToolStripStatusLabel.Text = "Loading items from " & Me.sentOnlyToMe.FolderPath
        For Each curItem In Me.sentOnlyToMe.Items
            Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
            If Not threadMap.ContainsKey(conversationIdx) Then

                Dim thread As New Thread(conversationIdx, 2)
                threadMap.Add(conversationIdx, thread)
                ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
            End If
        Next

        ToolStripStatusLabel.Text = "Loading items from " & Me.vipFolder.FolderPath
        For Each curItem In Me.vipFolder.Items
            Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
            If Not threadMap.ContainsKey(conversationIdx) Then

                Dim thread As New Thread(conversationIdx, 3)
                threadMap.Add(conversationIdx, thread)
                ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
            End If
        Next

        ToolStripStatusLabel.Text = "Loading items from " & Me.sentToMe.FolderPath
        For Each curItem In Me.sentToMe.Items
            Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
            If Not threadMap.ContainsKey(conversationIdx) Then

                Dim thread As New Thread(conversationIdx, 4)
                threadMap.Add(conversationIdx, thread)
                ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
            End If
        Next

        ToolStripStatusLabel.Text = "Loading items from " & Me.sentToMyGroup.FolderPath
        For Each curItem In Me.sentToMyGroup.Items
            Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
            If Not threadMap.ContainsKey(conversationIdx) Then

                Dim thread As New Thread(conversationIdx, 5)
                threadMap.Add(conversationIdx, thread)
                ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
            End If
        Next

        'ToolStripStatusLabel.Text = "Loading items from " & Me.verifyingFolder.FolderPath
        'For Each curItem In Me.verifyingFolder.Items
        '    Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
        '    If Not threadMap.ContainsKey(conversationIdx) Then

        '        Dim thread As New Thread(conversationIdx, 6)
        '        threadMap.Add(conversationIdx, thread)
        '        ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
        '    End If
        'Next

        'ToolStripStatusLabel.Text = "Loading items from " & Me.activeFolder.FolderPath
        'For Each curItem In Me.activeFolder.Items
        '    Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
        '    If Not threadMap.ContainsKey(conversationIdx) Then

        '        Dim thread As New Thread(conversationIdx, 7)
        '        threadMap.Add(conversationIdx, thread)
        '        ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
        '    End If
        'Next

        'ToolStripStatusLabel.Text = "Loading items from " & Me.backlogFolder.FolderPath
        'For Each curItem In Me.backlogFolder.Items
        '    Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
        '    If Not threadMap.ContainsKey(conversationIdx) Then

        '        Dim thread As New Thread(conversationIdx, 8)
        '        threadMap.Add(conversationIdx, thread)
        '        ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
        '    End If
        'Next

        'ToolStripStatusLabel.Text = "Loading items from " & Me.newFolder.FolderPath
        'For Each curItem In Me.newFolder.Items
        '    Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)
        '    If Not threadMap.ContainsKey(conversationIdx) Then

        '        Dim thread As New Thread(conversationIdx, 9)
        '        threadMap.Add(conversationIdx, thread)
        '        ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
        '    End If
        'Next

        'ToolStripStatusLabel.Text = "Loading items from " & Me.forFollowUpFolder.FolderPath
        'For Each curItem In Me.forFollowUpFolder.Items
        '    Dim curMail As Outlook.MailItem = TryCast(curItem, Outlook.MailItem)
        '    If curMail IsNot Nothing Then
        '        Dim conversationIdx = curMail.ConversationIndex.Substring(0, 44)
        '        Dim thread As Thread = Nothing

        '        If Not threadMap.TryGetValue(conversationIdx, thread) Then

        '            thread = New Thread(conversationIdx, 98)
        '            threadMap.Add(conversationIdx, thread)
        '            ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
        '        End If

        '        thread.AddEmail(New Email(curMail.EntryID, curMail.SentOn))
        '        EmailToolStripStatusLabel.Text = "Email " & (Me.mailIdx + 1) & " of " & thread.Emails.Count
        '    End If
        'Next

        Dim inboxFolder As Outlook.Folder = Me.olNs.DefaultStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox)
        ToolStripStatusLabel.Text = "Loading items from " & inboxFolder.FolderPath
        For Each curItem In inboxFolder.Items

            If curItem.MessageClass.ToString.StartsWith("IPM.Schedule.Meeting.") OrElse curItem.MessageClass.ToString.Equals("REPORT.IPM.Note.NDR") OrElse curItem.MessageClass.ToString.Equals("IPM.Note.Rules.OofTemplate.Microsoft") Then

                'Dim conversationIdx = curItem.ConversationIndex.Substring(0, 44)

                'Dim thread As Thread = Nothing

                'If threadMap.TryGetValue(conversationIdx, thread) Then

                '    thread.Priority = 1

                'Else

                '    thread = New Thread(conversationIdx, 1)
                '    threadMap.Add(conversationIdx, thread)
                '    ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count

                'End If

                olNs.GetItemFromID(curItem.EntryID).Display()

            Else

                Dim curMail As Outlook.MailItem = TryCast(curItem, Outlook.MailItem)
                If curMail IsNot Nothing AndAlso curMail.UnRead Then
                    Dim conversationIdx = curMail.ConversationIndex.Substring(0, 44)
                    Dim thread As Thread = Nothing

                    If Not threadMap.TryGetValue(conversationIdx, thread) Then

                        thread = New Thread(conversationIdx, 99)
                        threadMap.Add(conversationIdx, thread)
                        ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & threadMap.Count
                    End If

                    If curMail.Importance = OlImportance.olImportanceHigh Then
                        thread.Priority = 1
                    End If

                    thread.AddEmail(New Email(curMail.EntryID, curMail.SentOn))
                    EmailToolStripStatusLabel.Text = "Email " & (Me.mailIdx + 1) & " of " & thread.Emails.Count
                End If
            End If

        Next

        ' Retrieve all Threads from the dictionary into a list
        Me.threads = threadMap.Values.ToList()

        ' Sort the list by Priority ascending and StartDate
        Me.threads = Me.threads.OrderBy(Function(t) t.Priority).ThenBy(Function(t) t.StartDate).ToList()

    End Sub

    Sub DisplayEmail()
        Dim thread As Thread
        If threadIdx < Me.threads.Count Then
            thread = Me.threads.Item(Me.threadIdx)
            While Me.mailIdx >= thread.Emails.Count
                Me.mailIdx = 0
                Me.threadIdx += 1
                If threadIdx < Me.threads.Count Then
                    thread = Me.threads.Item(Me.threadIdx)
                End If
            End While


            ThreadToolStripStatusLabel.Text = "Thread " & (Me.threadIdx + 1) & " of " & Me.threads.Count & " P" & thread.Priority

            Dim email As Email = thread.Emails.Item(Me.mailIdx)

            EmailToolStripStatusLabel.Text = "Email " & (Me.mailIdx + 1) & " of " & thread.Emails.Count

            Dim itemEntryId As String = email.EntryId

            Me.dgMail = Me.olNs.GetItemFromID(itemEntryId)


            'Me.LogTextBox.AppendText(Now() & "> " & Me.dgMail.SentOn & " | " & Me.dgMail.SenderName & " | " & Me.dgMail.Subject & vbNewLine)
            Dim row0 As String() = {itemEntryId, Now(), Me.dgMail.SentOn, Me.dgMail.SenderName, Me.dgMail.Subject}

            If Me.LogDataGridView.InvokeRequired Then
                Me.LogDataGridView.Invoke(Sub()
                                              Me.LogDataGridView.Rows.Add(row0)
                                          End Sub)
            Else
                Me.LogDataGridView.Rows.Add(row0)
            End If

            Me.dgMail.Display()
            Me.mailIdx += 1

        End If

    End Sub

    Sub dgMail_Close(ByRef Cancel As Boolean) Handles dgMail.Close
        dgMail = Nothing
        Console.WriteLine("mail closed")
        DisplayEmail()
    End Sub

    Private Sub RefreshButton_Click(sender As Object, e As EventArgs) Handles RefreshButton.Click
        If dgMail IsNot Nothing Then
            Dim displayMail As Outlook.MailItem = dgMail
            dgMail = Nothing
            displayMail.UnRead = vbTrue
            displayMail.Close(OlInspectorClose.olDiscard)
        End If

        OrganizeEmail()
        ToolStripStatusLabel.Text = "Reading"
        DisplayEmail()

    End Sub

    Private Sub SkipThreadButton_Click(sender As Object, e As EventArgs) Handles SkipThreadButton.Click

        'Me.LogTextBox.AppendText(Now() & "> SKIP" & vbNewLine)
        Dim thread As Thread = Me.threads.Item(Me.threadIdx)

        For idxEmail = Me.mailIdx To thread.Emails.Count - 1 Step 1
            Dim email As Email = thread.Emails.Item(idxEmail)
            Dim itemEntryId As String = email.EntryId

            Dim mailItem As Outlook.MailItem = Me.olNs.GetItemFromID(itemEntryId)
            mailItem.UnRead = False

        Next idxEmail

        Me.mailIdx = 0
        Me.threadIdx += 1

        Me.dgMail.Close(OlInspectorClose.olDiscard)

    End Sub

    Private Sub LogDataGridView_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles LogDataGridView.CellDoubleClick
        If e.RowIndex >= 0 Then
            Dim entryId As String = LogDataGridView.Rows(e.RowIndex).Cells(0).Value.ToString()
            Dim mailItem As Outlook.MailItem = Me.olNs.GetItemFromID(entryId)
            mailItem.Display()
        End If
    End Sub

End Class
