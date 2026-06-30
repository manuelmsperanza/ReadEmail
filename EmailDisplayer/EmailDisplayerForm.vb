Imports System.ComponentModel
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Menu
Imports Microsoft.Office.Interop.Outlook

Public Class EmailDisplayerForm

    WithEvents olApp As Outlook.Application  'Riferimento processo OUTLOOK
    Private olNs As Outlook.NameSpace     'Namespace utilizzato per navigare i pst, ottenere la sessione, gli elementi selezionati...

    WithEvents dgMail As Outlook.MailItem

    Private sentOnlyToMe As Outlook.Folder
    Private sentToMe As Outlook.Folder
    Private vipFolder As Outlook.Folder
    Private vipOutFolder As Outlook.Folder
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

            If curFolder.FolderPath.Equals(searchFolderPrefix & "VIP out") Then
                Me.vipOutFolder = curFolder
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

        ToolStripStatusLabel.Text = "Loading items from " & Me.vipOutFolder.FolderPath
        For Each curItem In Me.vipOutFolder.Items
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

    Sub addDataGridRow(ByRef mailItem As Outlook.MailItem)
        'Me.LogTextBox.AppendText(Now() & "> " & Me.dgMail.SentOn & " | " & Me.dgMail.SenderName & " | " & Me.dgMail.Subject & vbNewLine)
        Dim row0 As String() = {mailItem.EntryID, Now(), mailItem.SentOn, mailItem.SenderName, mailItem.ConversationTopic}

        If Me.LogDataGridView.InvokeRequired Then
            Me.LogDataGridView.Invoke(Sub()
                                          Me.LogDataGridView.Rows.Add(row0)
                                          'Me.LogDataGridView.CurrentCell = Me.LogDataGridView.Rows(Me.LogDataGridView.Rows.Count - 1).Cells(0)
                                      End Sub)
        Else
            Me.LogDataGridView.Rows.Add(row0)
            'Me.LogDataGridView.CurrentCell = Me.LogDataGridView.Rows(Me.LogDataGridView.Rows.Count - 1).Cells(0)
        End If

        Console.WriteLine("exit addDataGridRow")
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
            Me.addDataGridRow(Me.dgMail)

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
            addDataGridRow(mailItem)
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

    Private Sub EmailDisplayerForm_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        If dgMail IsNot Nothing Then
            Dim displayMail As Outlook.MailItem = dgMail
            dgMail = Nothing
            displayMail.UnRead = vbTrue
        End If

    End Sub

    Private Sub olApp_NewMailEx(EntryIDCollection As String) Handles olApp.NewMailEx
        ' EntryIDCollection may contain multiple EntryIDs separated by commas
        Console.WriteLine("New mail received: " & EntryIDCollection)
        Dim entryIDs() As String = EntryIDCollection.Split(","c)
        For Each entryID In entryIDs


            Dim mailItem As Object = Me.olNs.GetItemFromID(entryID)

            Console.WriteLine(mailItem.MessageClass.ToString)
            Console.WriteLine(mailItem.Parent.ToString)

            'Try
            '    Dim mailItem As Outlook.MailItem = TryCast(Me.olNs.GetItemFromID(entryID), Outlook.MailItem)
            '    If mailItem IsNot Nothing Then
            '        ' Do something with the new mail, e.g. add to grid or display
            '        addDataGridRow(mailItem)
            '        ' Optionally, display the mail
            '        ' mailItem.Display()
            '    End If
            'Catch ex As Exception
            '    ' Handle exceptions (e.g., item is not a MailItem)
            'End Try
        Next
    End Sub

    Private Sub ExportMailButton_Click(sender As Object, e As EventArgs) Handles ExportMailButton.Click
        If LogDataGridView.SelectedRows.Count = 0 Then
            MessageBox.Show("Select one or more email rows to export.", "Export email", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Using exportDirectoryDialog As New FolderBrowserDialog()
            exportDirectoryDialog.Description = "Choose the directory where the selected emails and attachments will be saved."
            exportDirectoryDialog.ShowNewFolderButton = True

            If exportDirectoryDialog.ShowDialog(Me) <> DialogResult.OK Then
                Return
            End If

            For Each selectedRow As DataGridViewRow In LogDataGridView.SelectedRows
                If selectedRow.IsNewRow OrElse selectedRow.Cells(0).Value Is Nothing Then
                    Continue For
                End If

                Dim entryId As String = selectedRow.Cells(0).Value.ToString()
                Dim mailItem As Outlook.MailItem = TryCast(Me.olNs.GetItemFromID(entryId), Outlook.MailItem)

                If mailItem Is Nothing Then
                    Continue For
                End If

                ExportMailItem(mailItem, exportDirectoryDialog.SelectedPath)
            Next selectedRow
        End Using
    End Sub

    Private Sub ExportMailItem(mailItem As Outlook.MailItem, exportDirectory As String)
        Dim fileBaseName As String = GetMailFileBaseName(mailItem)
        Dim mailFilePath As String = GetUniqueFilePath(exportDirectory, fileBaseName & ".txt")

        mailItem.SaveAs(mailFilePath, OlSaveAsType.olTXT)

        For attachmentIndex As Integer = 1 To mailItem.Attachments.Count
            Dim attachment As Outlook.Attachment = mailItem.Attachments.Item(attachmentIndex)
            Dim attachmentFileName As String = GetUniqueFilePath(exportDirectory, fileBaseName & "_attachment_" & attachmentIndex.ToString("00") & "_" & SanitizeFileName(attachment.FileName))

            attachment.SaveAsFile(attachmentFileName)
        Next attachmentIndex
    End Sub

    Private Function GetMailFileBaseName(mailItem As Outlook.MailItem) As String
        Dim shortEntryId As String = mailItem.EntryID

        If shortEntryId.Length > 8 Then
            shortEntryId = shortEntryId.Substring(shortEntryId.Length - 8)
        End If

        Return SanitizeFileName(mailItem.SentOn.ToString("yyyyMMdd_HHmmss") & "_" & mailItem.Subject & "_" & shortEntryId)
    End Function

    Private Function SanitizeFileName(fileName As String) As String
        If String.IsNullOrWhiteSpace(fileName) Then
            Return "email"
        End If

        Dim sanitized As String = fileName

        For Each invalidCharacter As Char In Path.GetInvalidFileNameChars()
            sanitized = sanitized.Replace(invalidCharacter, "_"c)
        Next invalidCharacter

        sanitized = sanitized.Trim()

        If sanitized.Length > 120 Then
            sanitized = sanitized.Substring(0, 120)
        End If

        Return sanitized
    End Function

    Private Function GetUniqueFilePath(directoryPath As String, fileName As String) As String
        Dim fullPath As String = Path.Combine(directoryPath, fileName)

        If Not File.Exists(fullPath) Then
            Return fullPath
        End If

        Dim fileNameWithoutExtension As String = Path.GetFileNameWithoutExtension(fileName)
        Dim extension As String = Path.GetExtension(fileName)
        Dim fileIndex As Integer = 1

        Do
            fullPath = Path.Combine(directoryPath, fileNameWithoutExtension & "_" & fileIndex.ToString() & extension)
            fileIndex += 1
        Loop While File.Exists(fullPath)

        Return fullPath
    End Function
End Class
