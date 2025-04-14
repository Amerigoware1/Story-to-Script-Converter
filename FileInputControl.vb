Imports System.IO

Public Class FileInputControl '(WizardSteps(0)
    Public FilePath As String

    Public ReadOnly Property TextBoxContent As RichTextBox
        Get
            Return RTB1
        End Get
    End Property

    Private Sub FileInputControl_Load(sender As Object, e As EventArgs) Handles Me.Load
        If FilePath <> "" Then
            LoadFile(FilePath)
        End If
    End Sub

    Private Sub SelectFileButton_Click(sender As Object, e As EventArgs) Handles SelectFileButton.Click
        OFD.FileName = My.Settings.LastFilePath
        If OFD.ShowDialog() = DialogResult.OK Then
            FilePath = OFD.FileName
            LoadFile(FilePath)
        End If
    End Sub

    Public Sub LoadFile(filePath As String)
        RTB1.Clear()
        Try
            If Not String.IsNullOrEmpty(filePath) Then
                Dim fileExtension As String = IO.Path.GetExtension(filePath).ToLower()

                'Check for unsupported file types.
                If Not ({".txt", ".rtf", ".stc"}.Contains(fileExtension)) Then
                    MessageBox.Show("Unsupported file type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Exit Try
                End If

                Select Case fileExtension
                    Case ".rtf"
                        RTB1.LoadFile(filePath)
                    Case ".txt"
                        RTB1.Text = IO.File.ReadAllText(filePath)
                    Case ".stc"
                        Form1.CurrentStep = 3
                        Form1.BeginInvoke(Sub()
                                              Dim reviewControl As ReviewControl = DirectCast(Form1.WizardSteps(4), ReviewControl)
                                              reviewControl.TextToAnalyze = filePath
                                              reviewControl.NarratorVoice = "Cortana" ' Set a default value or retrieve from settings
                                          End Sub)
                        Form1.NextButton_Click(Nothing, Nothing)
                        Return ' Exit LoadFile since ReviewControl will load the file
                    Case Else
                        MessageBox.Show("Unsupported file type.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Select
                Form1.Text = "Story to Script Converter - " & IO.Path.GetFileName(filePath)
                FilePathLabel.Text = filePath
                Form1.NextButton.Enabled = True
                Form1.NextButton.Focus()
                My.Settings.LastFilePath = filePath
                My.Settings.Save()
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading file: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PasteButton_Click(sender As Object, e As EventArgs) Handles PasteButton.Click
        RTB1.Clear()
        RTB1.Paste()
        Form1.NextButton.Enabled = True
        Form1.NextButton.Focus()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        RTB1.Clear()
    End Sub

    Private Sub RTB1_DragEnter(sender As Object, e As DragEventArgs) Handles RTB1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub RTB1_DragDrop(sender As Object, e As DragEventArgs) Handles RTB1.DragDrop
        Try
            If e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim files As String() = CType(e.Data.GetData(DataFormats.FileDrop), String())
                If files.Length > 0 Then
                    FilePath = files(0)
                    If File.Exists(FilePath) Then
                        LoadFile(FilePath)
                        e.Effect = DragDropEffects.None
                    Else
                        MessageBox.Show("File not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If
                Else
                    MessageBox.Show("No files were dropped.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error during drag-and-drop: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
