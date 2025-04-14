<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class DialogueAssignmentControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.DialogueListView = New System.Windows.Forms.ListView()
        Me.Dialogs = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'DialogueListView
        '
        Me.DialogueListView.AutoArrange = False
        Me.DialogueListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Dialogs})
        Me.DialogueListView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DialogueListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.DialogueListView.HideSelection = False
        Me.DialogueListView.LabelWrap = False
        Me.DialogueListView.Location = New System.Drawing.Point(0, 0)
        Me.DialogueListView.MultiSelect = False
        Me.DialogueListView.Name = "DialogueListView"
        Me.DialogueListView.Size = New System.Drawing.Size(518, 362)
        Me.DialogueListView.TabIndex = 0
        Me.DialogueListView.UseCompatibleStateImageBehavior = False
        Me.DialogueListView.View = System.Windows.Forms.View.Details
        '
        'Dialogs
        '
        Me.Dialogs.Text = "Dialogs"
        Me.Dialogs.Width = 400
        '
        'DialogueAssignmentControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 14.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.DialogueListView)
        Me.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "DialogueAssignmentControl"
        Me.Size = New System.Drawing.Size(518, 362)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DialogueListView As ListView
    Friend WithEvents Dialogs As ColumnHeader
End Class
