<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FileInputControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.RTB1 = New System.Windows.Forms.RichTextBox()
        Me.SelectFileButton = New System.Windows.Forms.Button()
        Me.OFD = New System.Windows.Forms.OpenFileDialog()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.PasteButton = New System.Windows.Forms.Button()
        Me.FilePathLabel = New System.Windows.Forms.Label()
        Me.ClearButton = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'RTB1
        '
        Me.RTB1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RTB1.EnableAutoDragDrop = True
        Me.RTB1.Location = New System.Drawing.Point(3, 50)
        Me.RTB1.Name = "RTB1"
        Me.RTB1.Size = New System.Drawing.Size(512, 283)
        Me.RTB1.TabIndex = 0
        Me.RTB1.Text = ""
        '
        'SelectFileButton
        '
        Me.SelectFileButton.Location = New System.Drawing.Point(3, 21)
        Me.SelectFileButton.Name = "SelectFileButton"
        Me.SelectFileButton.Size = New System.Drawing.Size(75, 23)
        Me.SelectFileButton.TabIndex = 1
        Me.SelectFileButton.Text = "Select File"
        Me.SelectFileButton.UseVisualStyleBackColor = True
        '
        'OFD
        '
        Me.OFD.Filter = "All Text Types|*.txt;*.rtf;*.doc;*.docx;*.otf;*.stc"
        Me.OFD.Title = "New Story"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(85, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(19, 14)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Or"
        '
        'PasteButton
        '
        Me.PasteButton.Location = New System.Drawing.Point(110, 21)
        Me.PasteButton.Name = "PasteButton"
        Me.PasteButton.Size = New System.Drawing.Size(75, 23)
        Me.PasteButton.TabIndex = 3
        Me.PasteButton.Text = "Paste text"
        Me.PasteButton.UseVisualStyleBackColor = True
        '
        'FilePathLabel
        '
        Me.FilePathLabel.AutoSize = True
        Me.FilePathLabel.Location = New System.Drawing.Point(4, 4)
        Me.FilePathLabel.Name = "FilePathLabel"
        Me.FilePathLabel.Size = New System.Drawing.Size(0, 14)
        Me.FilePathLabel.TabIndex = 4
        '
        'ClearButton
        '
        Me.ClearButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClearButton.Location = New System.Drawing.Point(440, 21)
        Me.ClearButton.Name = "ClearButton"
        Me.ClearButton.Size = New System.Drawing.Size(75, 23)
        Me.ClearButton.TabIndex = 5
        Me.ClearButton.Text = "Clear"
        Me.ClearButton.UseVisualStyleBackColor = True
        '
        'FileInputControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(96.0!, 96.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi
        Me.Controls.Add(Me.ClearButton)
        Me.Controls.Add(Me.FilePathLabel)
        Me.Controls.Add(Me.PasteButton)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.SelectFileButton)
        Me.Controls.Add(Me.RTB1)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "FileInputControl"
        Me.Size = New System.Drawing.Size(518, 336)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents SelectFileButton As Button
    Friend WithEvents OFD As OpenFileDialog
    Friend WithEvents Label1 As Label
    Friend WithEvents PasteButton As Button
    Friend WithEvents FilePathLabel As Label
    Friend WithEvents ClearButton As Button
    Public WithEvents RTB1 As RichTextBox
End Class
