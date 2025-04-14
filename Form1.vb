'todo: Highlight spoken words in the ReviewControl
'todo:  Ensure that voice assignments can be modified after loading a .stc file (Clicking on PreviousButton)
'todo:  Add Aliases to characters in voice assignment control
'todo: Add fix for compount names like "Ellie Mae" in NameIdentificationControl
'todo: Save additional character information (age, gender) to the .stc file
'todo: fix non-names in NameIdentificationControl by comparing to namelist text file

Imports System.IO

Public Class Form1
    Private Args() As String = Environment.GetCommandLineArgs()
    Public CurrentStep As Integer = 0
    Public WizardSteps As UserControl() = {New FileInputControl(), New NameIdentificationControl(), New DialogueAssignmentControl, New VoiceAssignmentControl, New ReviewControl}
    Public Log As String = My.Application.Info.DirectoryPath & "\log.txt"

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If My.Settings.Location <> New Point(0, 0) Then
            Me.Location = My.Settings.Location
        End If
        If My.Settings.Size <> New Size(0, 0) Then
            Me.Size = My.Settings.Size
        End If
        Dim pProcess() As Process = System.Diagnostics.Process.GetProcessesByName("notepad")

        For Each p As Process In pProcess
            p.Kill()
        Next
        If Not File.Exists(Log) Then
            File.Create(Log).Close()
        Else
            File.WriteAllText(Log, "") 'clear the log file
        End If
        ' Initialize the wizard
        Panel1.Controls.Add(WizardSteps(0)) ' Use a more descriptive name
        WizardSteps(0).Dock = DockStyle.Fill

        ' Handle command-line arguments (with error handling)
        Try
            If Args.Length > 1 Then
                Dim FilePath As String = Args(1)
                If Not String.IsNullOrEmpty(filePath) Then
                    DirectCast(WizardSteps(0), FileInputControl).FilePath = FilePath
                    ' Check if the file is .stc
                    'If System.IO.Path.GetExtension(filePath).ToLower() = ".stc" Then
                    '    BeginInvoke(Sub()
                    '                    Dim reviewControl As ReviewControl = DirectCast(WizardSteps(4), ReviewControl)
                    '                    reviewControl.TextToAnalyze = FilePath
                    '                    reviewControl.NarratorVoice = "Cortana" ' Set a default value or retrieve from settings
                    '                End Sub)
                    '    NextButton_Click(Nothing, Nothing)
                    '    Return ' Exit LoadFile since ReviewControl will load the file
                    'End If
                End If
            End If
        Catch ex As Exception
            MessageBox.Show("Error processing command-line argument: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Public Sub NextButton_Click(sender As Object, e As EventArgs) Handles NextButton.Click
        If CurrentStep > 0 Then
            PreviousButton.Enabled = True
        End If

        If CurrentStep = 0 Then ' Assuming FileInputControl is step 0
            Dim fileInputControl As FileInputControl = TryCast(WizardSteps(CurrentStep), FileInputControl)
            If fileInputControl IsNot Nothing Then
                Dim text As String = fileInputControl.TextBoxContent.Text
                'Dim Fi As New FileInfo(fileInputControl.FilePath)
                'If Fi.Extension.ToLower() = ".stc" Then
                '    DirectCast(WizardSteps(4), ReviewControl).ReviewTxt.Text = text
                '    CurrentStep = 4
                'Else
                DirectCast(WizardSteps(CurrentStep + 1), NameIdentificationControl).TextToAnalyze = text
                'End If

                ' Pass text to the next step

            End If
            'NextButton.Enabled = False
        End If

        If CurrentStep = 1 Then ' Assuming NameIdentificationControl is step 1
            Dim nameControl As NameIdentificationControl = TryCast(WizardSteps(CurrentStep), NameIdentificationControl)
            If nameControl IsNot Nothing Then
                Dim selectedNames As List(Of String) = nameControl.GetSelectedNames()
                ' Pass selectedNames to the next step
                Dim dialogueControl As DialogueAssignmentControl = DirectCast(WizardSteps(CurrentStep + 1), DialogueAssignmentControl)
                dialogueControl.CharacterNames = selectedNames
                dialogueControl.PopulateCharacterComboBox() ' Call the new subroutine
            End If
            'NextButton.Enabled = False
        End If

        If CurrentStep = 2 Then ' Assuming DialogueAssignmentControl is step 2
            Dim dialogueControl As DialogueAssignmentControl = TryCast(WizardSteps(CurrentStep), DialogueAssignmentControl)
            If dialogueControl IsNot Nothing Then
                Dim assignments As Dictionary(Of String, List(Of String)) = dialogueControl.GetCharacterDialogueAssignments()
                ' Pass assignments to the next step

                DirectCast(WizardSteps(CurrentStep + 1), VoiceAssignmentControl).DialogueAssignments = assignments
            End If
            'NextButton.Enabled = False
        End If

        If CurrentStep = 3 Then ' Assuming VoiceAssignmentControl is step 3
            Dim voiceControl As VoiceAssignmentControl = TryCast(WizardSteps(CurrentStep), VoiceAssignmentControl)
            Dim FileInput As FileInputControl = TryCast(WizardSteps(0), FileInputControl)
            If voiceControl IsNot Nothing Then
                Dim characterVoiceAssignments As Dictionary(Of String, String) = voiceControl.GetCharacterVoiceAssignments()
                Dim dialogueAssignments As Dictionary(Of String, List(Of String)) = voiceControl.DialogueAssignments
                Dim reviewControl As ReviewControl = DirectCast(WizardSteps(CurrentStep + 1), ReviewControl)
                reviewControl.CharacterVoiceAssignments = characterVoiceAssignments
                reviewControl.DialogueAssignments = dialogueAssignments
                reviewControl.NarratorVoice = voiceControl.NarratorVoice
                ' Debugging: Inspect data after transfer
                Dim assignmentsString As String = ""
                For Each actor As String In reviewControl.CharacterVoiceAssignments.Keys
                    assignmentsString &= $"{actor}: {reviewControl.CharacterVoiceAssignments(actor)}{vbCrLf}"
                Next
                IO.File.AppendAllText(Log, $"Character Voice Assignments in ReviewControl:{vbCrLf}{assignmentsString}")

                Dim dialogueAssignmentsString As String = ""
                For Each kvp As KeyValuePair(Of String, List(Of String)) In reviewControl.DialogueAssignments
                    dialogueAssignmentsString &= $"{kvp.Key}: {String.Join(", ", kvp.Value.ToArray())}{vbCrLf}"
                Next
                IO.File.AppendAllText(Log, $"Dialogue Assignments in ReviewControl:{vbCrLf}{dialogueAssignmentsString}")

                IO.File.AppendAllText(Log, $"Text to Analyze in ReviewControl:{vbCrLf}{reviewControl.TextToAnalyze}")
                'end debugging

                reviewControl.TextToAnalyze = FileInput.TextBoxContent.Text
                reviewControl.PopulateVoicesComboBox()
            End If

            'NextButton.Enabled = False
        End If

        If CurrentStep < WizardSteps.Length - 1 Then
            Panel1.Controls.Clear()
            CurrentStep += 1
            Panel1.Controls.Add(WizardSteps(CurrentStep))
            WizardSteps(CurrentStep).Dock = DockStyle.Fill
        ElseIf CurrentStep = 4 Then
            NextButton.Enabled = False
        End If
    End Sub

    Private Sub PreviousButton_Click(sender As Object, e As EventArgs) Handles PreviousButton.Click
        If currentStep > 0 Then
            Panel1.Controls.Clear()
            currentStep -= 1
            Panel1.Controls.Add(WizardSteps(CurrentStep))
            WizardSteps(CurrentStep).Dock = DockStyle.Fill
        ElseIf CurrentStep <= 0 Then
            PreviousButton.Enabled = False
        End If
    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        My.Settings.Location = Me.Location
        My.Settings.Size = Me.Size
        My.Settings.Save()
    End Sub

End Class


