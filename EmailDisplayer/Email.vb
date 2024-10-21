Public Class Email

    Public Property EntryId As String
    Public Property SentOn As Date

    Public Sub New(entryId As String, sentOn As Date)
        Me.EntryId = entryId
        Me.SentOn = sentOn
    End Sub

End Class
