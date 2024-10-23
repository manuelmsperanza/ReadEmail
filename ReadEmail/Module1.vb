Imports Microsoft.Office.Interop.Outlook

Module Module1


    Sub Main()

        Dim olApp As Outlook.Application  'Riferimento processo OUTLOOK
        Dim olNs As Outlook.NameSpace     'Namespace utilizzato per navigare i pst, ottenere la sessione, gli elementi selezionati...

        olApp = New Outlook.Application
        olNs = olApp.GetNamespace("MAPI")
        Dim defaultStore As Outlook.Store = olNs.DefaultStore
        Dim inboxFolder As Outlook.Folder = defaultStore.GetDefaultFolder(OlDefaultFolders.olFolderInbox)

        For Each curItem In inboxFolder.Items


            If curItem.MessageClass.ToString.StartsWith("IPM.Schedule.Meeting.") AndAlso curItem.UnRead Then

                Console.WriteLine(curItem.MessageClass & " - " & curItem.Subject)

                olNs.GetItemFromID(curItem.EntryID).Display()
                Console.WriteLine("Press enter to continue")
                Console.ReadLine()
            ElseIf Not curItem.MessageClass.Equals("IPM.Note") Then
                Console.WriteLine(curItem.MessageClass & " - " & curItem.Subject)
            End If
        Next

        Console.WriteLine("Press enter to end")
        Console.ReadLine()

    End Sub




End Module
