Module Module1

    Function GetFolderByPath(ByRef rootFolder As Outlook.Folder, ByVal folderPath As String) As Outlook.Folder
        Dim tempFolder As Outlook.Folder = rootFolder
        Dim i As Integer

        'On Error GoTo GetFolder_Error
        If InStr(folderPath, "\\") = 1 Then
            folderPath = Left(folderPath, Len(folderPath) - 2)
        End If

        'Convert folderpath to array
        Dim foldersArray = Split(folderPath, "\")

        'Set TestFolder = Application.Session.Folders.item(FoldersArray(0))
        If Not tempFolder Is Nothing Then
            For i = 0 To UBound(foldersArray, 1)
                Dim subFolders As Outlook.Folders = tempFolder.Folders
                tempFolder = subFolders.Item(foldersArray(i))
                If tempFolder Is Nothing Then
                    tempFolder = subFolders.Add(foldersArray(i))
                End If
            Next
        End If

        'Return the TestFolder
        GetFolderByPath = tempFolder
        Exit Function

GetFolder_Error:
        GetFolderByPath = Nothing
        Exit Function
    End Function

    Sub Main()
        Dim olApp As Outlook.Application = New Outlook.Application
        Dim olNs As Outlook.NameSpace = olApp.GetNamespace("MAPI")
        Dim defaultStore As Outlook.Store = olNs.DefaultStore


        Dim targetStoreName As String
        targetStoreName = "\\Online Archive - " & defaultStore.DisplayName

        Dim sourceStoreRootFolder As Outlook.Folder = olNs.DefaultStore.GetRootFolder
        Dim targetStoreRootFolder As Outlook.Folder

        Dim folderList(19)
        folderList(0) = "IT - AOM Supporto Salesforce\Report"
        folderList(1) = "IT - AOM Supporto Salesforce\REQUEST PR"
        folderList(2) = "IT - AOM Supporto Salesforce\Spark Europe"
        folderList(3) = "IT - AOM Supporto Salesforce\NoOutbound_NoBI"
        folderList(4) = "IT - AOM Supporto Salesforce\@service.sky.it"
        folderList(5) = "IT - AOM Supporto Salesforce\@salesforce.com"
        folderList(6) = "IT - AOM Supporto Salesforce\@odaseva"
        folderList(7) = "IT - AOM Supporto Salesforce\Others"
        folderList(8) = "IT - AOM Supporto Salesforce\Others\ADV SalesForce"
        folderList(9) = "IT - AOM Supporto Salesforce\Others\Case update Batch"
        folderList(10) = "IT - AOM Supporto Salesforce\Others\Jtool"
        folderList(11) = "IT - AOM Supporto Salesforce\Others\Sandbox: Report"
        folderList(12) = "IT - AOM Supporto Salesforce\Others\Sandbox @salesforce.com"
        folderList(13) = "IT - AOM Supporto Salesforce\Bonifiche"
        folderList(14) = "IT - AOM Supporto Salesforce\Heroku"
        folderList(15) = "IT - AOM Supporto Salesforce\Flow Application"
        folderList(16) = "IT - AOM Supporto Salesforce\SFDC ApexApplication"
        folderList(17) = "IT - AOM Supporto Salesforce\@int.sky.it"
        folderList(18) = "IT - AOM Supporto Salesforce\IT-SA"


        For Each curStore In olNs.Stores


            Dim rootFolder As Outlook.Folder
            rootFolder = curStore.GetRootFolder

            Select Case rootFolder.FullFolderPath
                Case targetStoreName
                    targetStoreRootFolder = rootFolder
            End Select
        Next curStore

        For Each folderPath In folderList

            Dim sourceFolder As Outlook.Folder = GetFolderByPath(sourceStoreRootFolder, folderPath)

            If sourceFolder.Items.Count > 0 Then

                Console.WriteLine("Archiving " & folderPath & " [" & sourceFolder.Items.Count & "]")



                Dim targetFolder As Outlook.Folder = GetFolderByPath(targetStoreRootFolder, folderPath)

                For idxEmail = sourceFolder.Items.Count To 1 Step -1
                    Dim firstEmailItem As Outlook.MailItem = sourceFolder.Items.GetFirst
                    firstEmailItem.Move(targetFolder)
                Next idxEmail
            End If

        Next folderPath

    End Sub

End Module
