Public Class Thread

    Public Property ConversationId As String
    Public Property Priority As Integer
    Public Property StartDate As Date
    Public Property EntryId As String ' New property
    Private _emails As List(Of Email)

    ' Public read-only property to access the emails
    Public ReadOnly Property Emails As List(Of Email)
        Get
            Return _emails
        End Get
    End Property

    Public Sub New(conversationId As String, priority As Integer)
        Me.ConversationId = conversationId.Substring(0, 44)
        Me.Priority = priority
        Me._emails = New List(Of Email)()
        Me.StartDate = Date.MaxValue ' Initialize with a maximum date
        Me.EntryId = Nothing ' Initialize EntryId
    End Sub

    ' Method to add an email to the thread
    Public Sub AddEmail(email As Email)
        _emails.Add(email)
        ' Sort the emails by SentOn date
        _emails = _emails.OrderBy(Function(e) e.SentOn).ToList()
        ' Update the StartDate to the oldest SentOn date
        If email.SentOn < Me.StartDate Then
            Me.StartDate = email.SentOn
            ' Update EntryId to the EntryId of the oldest email
            Me.EntryId = email.EntryId
        End If
        ' If EntryId is not set yet (first email), set EntryId
        If String.IsNullOrEmpty(Me.EntryId) Then
            Me.EntryId = email.EntryId
        End If
    End Sub
End Class
