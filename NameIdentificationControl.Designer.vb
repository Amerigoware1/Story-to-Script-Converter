<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class NameIdentificationControl
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
        Me.PotentialNameListBox = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.StatusLbl = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'PotentialNameListBox
        '
        Me.PotentialNameListBox.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PotentialNameListBox.CheckOnClick = True
        Me.PotentialNameListBox.FormattingEnabled = True
        Me.PotentialNameListBox.Location = New System.Drawing.Point(3, 28)
        Me.PotentialNameListBox.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.PotentialNameListBox.Name = "PotentialNameListBox"
        Me.PotentialNameListBox.Size = New System.Drawing.Size(512, 304)
        Me.PotentialNameListBox.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(4, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(294, 14)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Check each actual name (with dialogue parts) in your story."
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ProgressBar1.Location = New System.Drawing.Point(0, 339)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(518, 23)
        Me.ProgressBar1.TabIndex = 2
        '
        'StatusLbl
        '
        Me.StatusLbl.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.StatusLbl.AutoSize = True
        Me.StatusLbl.BackColor = System.Drawing.Color.Transparent
        Me.StatusLbl.Location = New System.Drawing.Point(0, 344)
        Me.StatusLbl.Name = "StatusLbl"
        Me.StatusLbl.Size = New System.Drawing.Size(0, 14)
        Me.StatusLbl.TabIndex = 3
        '
        'NameIdentificationControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.StatusLbl)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PotentialNameListBox)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Margin = New System.Windows.Forms.Padding(3, 4, 3, 4)
        Me.Name = "NameIdentificationControl"
        Me.Size = New System.Drawing.Size(518, 362)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents PotentialNameListBox As CheckedListBox
    Friend WithEvents Label1 As Label
    Friend WithEvents ProgressBar1 As ProgressBar
    Friend WithEvents StatusLbl As Label
End Class
