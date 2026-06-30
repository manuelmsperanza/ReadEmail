<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class EmailDisplayerForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.ToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ThreadToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.EmailToolStripStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.RefreshButton = New System.Windows.Forms.Button()
        Me.SkipThreadButton = New System.Windows.Forms.Button()
        Me.LogDataGridView = New System.Windows.Forms.DataGridView()
        Me.MailId = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ReadOn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SentOn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.SenderName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Subject = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.ExportMailButton = New System.Windows.Forms.Button()
        Me.StatusStrip.SuspendLayout()
        CType(Me.LogDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'StatusStrip
        '
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ToolStripStatusLabel, Me.ThreadToolStripStatusLabel, Me.EmailToolStripStatusLabel})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 421)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Size = New System.Drawing.Size(1103, 22)
        Me.StatusStrip.TabIndex = 0
        Me.StatusStrip.Text = "StatusStrip1"
        '
        'ToolStripStatusLabel
        '
        Me.ToolStripStatusLabel.Name = "ToolStripStatusLabel"
        Me.ToolStripStatusLabel.Size = New System.Drawing.Size(39, 17)
        Me.ToolStripStatusLabel.Text = "Ready"
        '
        'ThreadToolStripStatusLabel
        '
        Me.ThreadToolStripStatusLabel.Name = "ThreadToolStripStatusLabel"
        Me.ThreadToolStripStatusLabel.Size = New System.Drawing.Size(78, 17)
        Me.ThreadToolStripStatusLabel.Text = "Thread # of #"
        '
        'EmailToolStripStatusLabel
        '
        Me.EmailToolStripStatusLabel.Name = "EmailToolStripStatusLabel"
        Me.EmailToolStripStatusLabel.Size = New System.Drawing.Size(70, 17)
        Me.EmailToolStripStatusLabel.Text = "Email # of #"
        '
        'RefreshButton
        '
        Me.RefreshButton.Location = New System.Drawing.Point(1016, 383)
        Me.RefreshButton.Name = "RefreshButton"
        Me.RefreshButton.Size = New System.Drawing.Size(75, 23)
        Me.RefreshButton.TabIndex = 1
        Me.RefreshButton.Text = "Refresh"
        Me.RefreshButton.UseVisualStyleBackColor = True
        '
        'SkipThreadButton
        '
        Me.SkipThreadButton.Location = New System.Drawing.Point(930, 383)
        Me.SkipThreadButton.Name = "SkipThreadButton"
        Me.SkipThreadButton.Size = New System.Drawing.Size(75, 23)
        Me.SkipThreadButton.TabIndex = 3
        Me.SkipThreadButton.Text = "Skip Thread"
        Me.SkipThreadButton.UseVisualStyleBackColor = True
        '
        'LogDataGridView
        '
        Me.LogDataGridView.AllowUserToAddRows = False
        Me.LogDataGridView.AllowUserToDeleteRows = False
        Me.LogDataGridView.AllowUserToOrderColumns = True
        Me.LogDataGridView.AllowUserToResizeRows = False
        Me.LogDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.LogDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.MailId, Me.ReadOn, Me.SentOn, Me.SenderName, Me.Subject})
        Me.LogDataGridView.Location = New System.Drawing.Point(12, 12)
        Me.LogDataGridView.Name = "LogDataGridView"
        Me.LogDataGridView.ReadOnly = True
        Me.LogDataGridView.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.LogDataGridView.Size = New System.Drawing.Size(1079, 364)
        Me.LogDataGridView.TabIndex = 4
        '
        'MailId
        '
        Me.MailId.HeaderText = "Mail ID"
        Me.MailId.Name = "MailId"
        Me.MailId.ReadOnly = True
        Me.MailId.Visible = False
        '
        'ReadOn
        '
        Me.ReadOn.HeaderText = "Read On"
        Me.ReadOn.Name = "ReadOn"
        Me.ReadOn.ReadOnly = True
        '
        'SentOn
        '
        Me.SentOn.HeaderText = "Sent On"
        Me.SentOn.Name = "SentOn"
        Me.SentOn.ReadOnly = True
        '
        'SenderName
        '
        Me.SenderName.HeaderText = "Sender"
        Me.SenderName.Name = "SenderName"
        Me.SenderName.ReadOnly = True
        '
        'Subject
        '
        Me.Subject.HeaderText = "Subject"
        Me.Subject.Name = "Subject"
        Me.Subject.ReadOnly = True
        '
        'ExportMailButton
        '
        Me.ExportMailButton.Location = New System.Drawing.Point(849, 383)
        Me.ExportMailButton.Name = "ExportMailButton"
        Me.ExportMailButton.Size = New System.Drawing.Size(75, 23)
        Me.ExportMailButton.TabIndex = 5
        Me.ExportMailButton.Text = "Export"
        Me.ExportMailButton.UseVisualStyleBackColor = True
        '
        'EmailDisplayerForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1103, 443)
        Me.Controls.Add(Me.ExportMailButton)
        Me.Controls.Add(Me.LogDataGridView)
        Me.Controls.Add(Me.SkipThreadButton)
        Me.Controls.Add(Me.RefreshButton)
        Me.Controls.Add(Me.StatusStrip)
        Me.Name = "EmailDisplayerForm"
        Me.Text = "EmailDisplayer"
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        CType(Me.LogDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents StatusStrip As StatusStrip
    Friend WithEvents ToolStripStatusLabel As ToolStripStatusLabel
    Friend WithEvents ThreadToolStripStatusLabel As ToolStripStatusLabel
    Friend WithEvents EmailToolStripStatusLabel As ToolStripStatusLabel
    Friend WithEvents RefreshButton As Button
    Friend WithEvents SkipThreadButton As Button
    Friend WithEvents LogDataGridView As DataGridView
    Friend WithEvents MailId As DataGridViewTextBoxColumn
    Friend WithEvents ReadOn As DataGridViewTextBoxColumn
    Friend WithEvents SentOn As DataGridViewTextBoxColumn
    Friend WithEvents SenderName As DataGridViewTextBoxColumn
    Friend WithEvents Subject As DataGridViewTextBoxColumn
    Friend WithEvents ExportMailButton As Button
End Class
