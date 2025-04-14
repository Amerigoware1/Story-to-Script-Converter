Imports System.IO
Imports System.Speech.Synthesis
Imports System.Text.RegularExpressions
Imports System.Web
Imports System.Xml
Imports System.Diagnostics

Public Class ReviewControl '(WizardSteps(4)
    Private WithEvents Vox As New SpeechSynthesizer
    Public Property CharacterVoiceAssignments As Dictionary(Of String, String)
    Public Property DialogueAssignments As Dictionary(Of String, List(Of String))
    Public Property NarratorLines As New List(Of String)
    Public Property NarratorVoice As String
    Public Property TextToAnalyze As String
    Private SpeechQueue As New Queue(Of String)()
    Private IsPaused As Boolean = False
    Private DefaultCharacterVoice As String = "Cortana"
    Private CurrentText As String = ""
    Private voiceCache As New Dictionary(Of String, SpeechSynthesizer) ' Cache for voice objects
    Private MyXMLFile As String
    Public FilePath As String
    Private currentReviewTextIndex As Integer = 0
    Private currentlySpeakingLine As String = ""
    Private isHighlighting As Boolean = False ' Flag to prevent overlapping highlights

    Private Sub ReviewControl_Load(sender As Object, e As EventArgs) Handles Me.Load
        If Form1.NextButton.Enabled = True Then
            Form1.NextButton.Enabled = False
        End If
        If NarratorVoice = "" Then
            NarratorVoice = "Cortana"
        End If
        Try
            VoicesCBO.SelectedItem = (NarratorVoice)

            ' Get the VoiceInfo object for the narrator voice
            Dim narratorVoiceInfo As VoiceInfo = Vox.GetInstalledVoices().FirstOrDefault(Function(v) v.VoiceInfo.Name = NarratorVoice)?.VoiceInfo

            If narratorVoiceInfo IsNot Nothing Then

                ' Initialize the voice cache with the narrator voice
                Dim narratorVox As New SpeechSynthesizer() ' Create a new SpeechSynthesizer
                narratorVox.SelectVoice(narratorVoiceInfo.Name) ' Select the narrator voice
                voiceCache.Add("NARRATOR", narratorVox) ' Add to the cache
            Else
                ' Handle the case where the narrator voice is not found
                IO.File.WriteAllText(Form1.Log, "Narrator voice not found." & vbNewLine)
            End If

        Catch ex As Exception
            IO.File.WriteAllText(Form1.Log, $"Error: {ex.Message}" & vbNewLine)
        End Try
        Application.DoEvents()

        If TextToAnalyze.EndsWith(".stc") Then
            LoadStcFile(TextToAnalyze)
        Else
            PopulateReviewText()
        End If

    End Sub

    Public Sub PopulateVoicesComboBox() ' New subroutine to populate VoicesCBO
        VoicesCBO.Items.AddRange(GetAvailableVoices())
    End Sub

    Private Function GetAvailableVoices() As String()
        Return Vox.GetInstalledVoices().
        Where(Function(v) v.Enabled AndAlso Not v.VoiceInfo.Name.StartsWith("IVONA", StringComparison.OrdinalIgnoreCase)).
        Select(Function(v) v.VoiceInfo.Name).
        ToArray()
    End Function

    'Private Sub PopulateReviewText()
    '    If CharacterVoiceAssignments IsNot Nothing AndAlso DialogueAssignments IsNot Nothing Then
    '        Dim reviewText As New System.Text.StringBuilder()
    '        Dim allLines As String() = TextToAnalyze.Split({vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
    '        Dim dialogueLines As New List(Of String)
    '        For Each kvp As KeyValuePair(Of String, List(Of String)) In DialogueAssignments
    '            dialogueLines.AddRange(kvp.Value)
    '        Next
    '        ' Debugging: Inspect DialogueAssignments
    '        Dim dialogueAssignmentsString As String = ""
    '        For Each kvp As KeyValuePair(Of String, List(Of String)) In DialogueAssignments
    '            dialogueAssignmentsString &= $"{kvp.Key}: {String.Join(", ", kvp.Value.ToArray())}{vbCrLf}"
    '        Next
    '        IO.File.AppendAllText(Form1.Log, $"Dialogue Assignments in ReviewControl:{vbCrLf}{dialogueAssignmentsString}Dialogue Assignments Debug" & vbNewLine)
    '        'Debugging: Inspect CharacterVoiceAssignments
    '        Dim CharacterVoiceAssignmentsString As String = ""
    '        For Each kvp As KeyValuePair(Of String, String) In CharacterVoiceAssignments
    '            CharacterVoiceAssignmentsString &= $"{kvp.Key}: {kvp.Value}{vbCrLf}"
    '        Next
    '        IO.File.AppendAllText(Form1.Log, $"CharacterVoiceAssignments in ReviewControl:{vbCrLf}{CharacterVoiceAssignmentsString}  CharacterVoiceAssignments Debug" & vbNewLine)
    '        For Each kvp As KeyValuePair(Of String, List(Of String)) In DialogueAssignments
    '            dialogueLines.AddRange(kvp.Value)
    '        Next
    '        For Each line As String In allLines
    '            Dim parts As String() = Regex.Split(line, "(""[^""]+"")") ' Split by quoted text
    '            For i As Integer = 0 To parts.Length - 1
    '                Dim part As String = parts(i).Trim()
    '                If Not String.IsNullOrEmpty(part) Then
    '                    If i Mod 2 = 1 Then ' Odd indices are dialogue lines
    '                        For Each character As String In CharacterVoiceAssignments.Keys
    '                            If DialogueAssignments.ContainsKey(character) AndAlso DialogueAssignments(character).Contains(part) Then
    '                                reviewText.AppendLine($"<{character}>{part}</{character}>")
    '                                Exit For
    '                            End If
    '                        Next
    '                    Else ' Even indices are narrator lines
    '                        reviewText.AppendLine($"<narrator>{part}</narrator>")
    '                    End If
    '                End If
    '            Next
    '        Next
    '        ' Debugging: Inspect reviewText
    '        IO.File.AppendAllText(Form1.Log, $"Review Text:{vbCrLf}{reviewText} & Review Text Debug" & vbCrLf)

    '        ReviewTxt.Text = reviewText.ToString()
    '    End If
    'End Sub

    Private Sub PopulateReviewText()
        If CharacterVoiceAssignments IsNot Nothing AndAlso DialogueAssignments IsNot Nothing Then
            Dim reviewText As New System.Text.StringBuilder()

            ' --- Optional: Log Input Data (Keep your existing logging if helpful) ---
            ' Log DialogueAssignments
            Dim dialogueAssignmentsString As New System.Text.StringBuilder()
            For Each kvp As KeyValuePair(Of String, List(Of String)) In DialogueAssignments
                dialogueAssignmentsString.AppendLine($"{kvp.Key}: {String.Join(" | ", kvp.Value)}") ' Using | as separator for clarity in log
            Next
            IO.File.AppendAllText(Form1.Log, $"DEBUG: Dialogue Assignments:{vbCrLf}{dialogueAssignmentsString}{vbCrLf}")

            ' Log CharacterVoiceAssignments
            Dim characterVoiceAssignmentsString As New System.Text.StringBuilder()
            For Each kvp As KeyValuePair(Of String, String) In CharacterVoiceAssignments
                characterVoiceAssignmentsString.AppendLine($"{kvp.Key}: {kvp.Value}")
            Next
            IO.File.AppendAllText(Form1.Log, $"DEBUG: Character Voice Assignments:{vbCrLf}{characterVoiceAssignmentsString}{vbCrLf}")
            ' --- End Optional Logging ---


            ' 1. Split the text into paragraphs/blocks based on blank lines
            '    Regex explanation: Matches one or more occurrences of:
            '    \r?\n : A line break (Windows or Unix style)
            '    \s* : Followed by zero or more whitespace characters (on the potentially blank line)
            '    \r?\n : Followed by another line break
            '    Using {2,} means matching 2 or more consecutive line breaks possibly separated by whitespace.
            Dim textBlocks As String() = Regex.Split(TextToAnalyze, "(\r?\n\s*){2,}")

            Dim isFirstBlock As Boolean = True ' To control initial spacing

            For Each block As String In textBlocks
                Dim currentBlock As String = block.Trim() ' Trim whitespace from the block itself

                If String.IsNullOrWhiteSpace(currentBlock) Then
                    ' This block was essentially a blank line separator.
                    ' We respect this separation but don't add extra blank lines yet,
                    ' let the processing of the *next* content block handle spacing if needed.
                    Continue For ' Skip processing purely whitespace blocks
                End If

                ' Add a blank line *before* processing a new non-empty block,
                ' *unless* it's the very first block.
                If Not isFirstBlock Then
                    reviewText.AppendLine() ' Add visual separation between paragraphs
                End If
                isFirstBlock = False ' No longer the first block

                ' 2. Process each non-empty block line by line (or as a whole if no internal breaks)
                '    Using StringSplitOptions.None here to potentially catch lines intended for dialogue
                '    that might have surrounding whitespace, relying on later Trim().
                Dim blockLines As String() = currentBlock.Split({vbCrLf}, StringSplitOptions.None)

                For Each line As String In blockLines
                    Dim trimmedLine = line.Trim() ' Trim each line *before* splitting dialogue/narration
                    If String.IsNullOrEmpty(trimmedLine) Then Continue For ' Skip empty lines within a block

                    ' 3. Split the line into dialogue and narration parts
                    '    Regex: (""[^""]+"") - Captures quoted strings
                    Dim parts As String() = Regex.Split(trimmedLine, "(""[^""]+"")")

                    For i As Integer = 0 To parts.Length - 1
                        Dim part As String = parts(i) ' Don't trim yet, dialogue lookup needs original quotes/spacing maybe

                        If String.IsNullOrWhiteSpace(part) Then Continue For ' Skip empty parts resulting from split

                        Dim trimmedPart = part.Trim() ' Trim for outputting narration

                        If i Mod 2 = 1 Then ' Odd indices are dialogue lines (quoted part)
                            ' Dialogue part still has quotes, e.g., ""Hello!""
                            Dim dialogueContent As String = part ' Keep original for lookup
                            Dim foundCharacter As Boolean = False

                            For Each character As String In CharacterVoiceAssignments.Keys
                                ' IMPORTANT: Check if DialogueAssignments stores keys WITH or WITHOUT quotes!
                                ' Assuming DialogueAssignments stores the dialogue *with* quotes:
                                If DialogueAssignments.ContainsKey(character) AndAlso DialogueAssignments(character).Contains(dialogueContent) Then
                                    reviewText.AppendLine($"<{character}>{dialogueContent}</{character}>")
                                    foundCharacter = True
                                    Exit For
                                End If
                            Next

                            If Not foundCharacter Then
                                ' Fallback: Dialogue found but not assigned to a known character.
                                ' Option 1: Assign to narrator
                                ' reviewText.AppendLine($"<narrator>{dialogueContent}</narrator>")
                                ' Option 2: Assign to a default voice
                                ' reviewText.AppendLine($"<Default>{dialogueContent}</Default>")
                                ' Option 3: Leave untagged (TTS might read it in narrator/default voice anyway)
                                ' reviewText.AppendLine(dialogueContent)
                                ' Option 4: Assign to narrator (Often safest if unattributed)
                                reviewText.AppendLine($"<narrator>{dialogueContent}</narrator>")
                                IO.File.AppendAllText(Form1.Log, $"WARNING: Unassigned dialogue found: {dialogueContent}{vbCrLf}")
                            End If
                        Else ' Even indices are narrator lines (non-quoted part)
                            ' Don't add narrator tags if the part is only whitespace
                            If Not String.IsNullOrWhiteSpace(trimmedPart) Then
                                reviewText.AppendLine($"<narrator>{trimmedPart}</narrator>")
                            End If
                        End If
                    Next ' Next part
                Next ' Next line in block
            Next ' Next block

            ' --- Final Output ---
            Dim finalOutput As String = reviewText.ToString()
            ' Optional: Remove trailing blank lines if the input ended with them
            finalOutput = Regex.Replace(finalOutput, "(\r?\n\s*)+$", "")

            IO.File.AppendAllText(Form1.Log, $"Review Text Output:{vbCrLf}{finalOutput}{vbCrLf}--- End Review Text Output ---{vbCrLf}")
            ReviewTxt.Text = finalOutput

        End If
    End Sub

    Private Sub ReadBtn_Click(sender As Object, e As EventArgs) Handles ReadBtn.Click
        ReadBtn.Enabled = False
        PauseBtn.Enabled = True
        StopBtn.Enabled = True
        Try
            If VoicesCBO.SelectedItem IsNot Nothing AndAlso Not String.IsNullOrEmpty(ReviewTxt.Text) Then
                SpeechQueue.Clear() ' Clear previous speech queue
                currentReviewTextIndex = 0 ' Reset the index
                currentlySpeakingLine = ""

                Dim lines As String() = ReviewTxt.Text.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

                For Each line As String In lines
                    Dim match As Match = Regex.Match(line, "^<([^>]+)>(.*?)</\1>$")

                    If match.Success Then
                        Dim tag As String = match.Groups(1).Value.Trim()
                        Dim dialogue As String = match.Groups(2).Value.Trim()

                        If tag.ToLower() = "narrator" Then
                            SpeechQueue.Enqueue("NARRATOR|" & dialogue)
                        ElseIf CharacterVoiceAssignments.ContainsKey(tag) Then
                            SpeechQueue.Enqueue(tag & "|" & dialogue)
                        Else
                            SpeechQueue.Enqueue("DEFAULT|" & dialogue)
                        End If
                    Else
                        SpeechQueue.Enqueue("NARRATOR|" & line.Trim()) ' Treat untagged text as narration
                    End If
                Next

                If SpeechQueue.Count > 0 Then
                    SpeakNextLine()
                End If
            Else
                MessageBox.Show("Please select a voice and ensure the review text is not empty.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If
        Catch ex As Exception
            MessageBox.Show($"Error speaking text: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SpeakNextLine()
        If SpeechQueue.Count > 0 AndAlso Not IsPaused Then
            Dim entry As String = SpeechQueue.Dequeue()
            Dim parts() As String = entry.Split("|"c)
            Dim voiceTag As String = parts(0)
            Dim textToSpeak As String = parts(1)
            currentlySpeakingLine = textToSpeak ' Store the line being spoken

            ' Find the starting index of the current line in ReviewTxt
            Dim indexInReviewTxt As Integer = ReviewTxt.Find(textToSpeak, currentReviewTextIndex, RichTextBoxFinds.MatchCase) ' Consider removing MatchCase if needed

            If indexInReviewTxt <> -1 Then
                currentReviewTextIndex = indexInReviewTxt
            Else
                ' Handle the case where the line is not found (shouldn't happen in normal circumstances)
                Debug.WriteLine($"Warning: Could not find line '{textToSpeak}' in ReviewTxt starting from index {currentReviewTextIndex}.")
            End If

            Select Case voiceTag.ToLower()
                Case "narrator"
                    Vox.SelectVoice(NarratorVoice)
                Case "default"
                    Vox.SelectVoice(DefaultCharacterVoice)
                Case Else
                    If CharacterVoiceAssignments.ContainsKey(voiceTag) Then
                        Vox.SelectVoice(CharacterVoiceAssignments(voiceTag))
                    Else
                        Vox.SelectVoice(DefaultCharacterVoice)
                    End If
            End Select
            Vox.SpeakAsync(textToSpeak) ' Speak asynchronously
        End If
    End Sub

    Private Sub Vox_SpeakCompleted(sender As Object, e As SpeakCompletedEventArgs) Handles Vox.SpeakCompleted
        If SpeechQueue.Count > 0 AndAlso Not IsPaused Then
            SpeakNextLine()
        Else
            ReadBtn.Enabled = True
            PauseBtn.Enabled = False
            StopBtn.Enabled = False
            ReviewTxt.Select(Nothing, Nothing)
        End If
    End Sub

    Private Sub PauseBtn_Click(sender As Object, e As EventArgs) Handles PauseBtn.Click
        If Vox.State = SynthesizerState.Speaking Then
            Vox.Pause()
            IsPaused = True
            PauseBtn.Text = "Resume"
        ElseIf Vox.State = SynthesizerState.Paused Then
            Vox.Resume()
            IsPaused = False
            SpeakNextLine()
            PauseBtn.Text = "Pause"
        End If
    End Sub

    Private Sub StopBtn_Click(sender As Object, e As EventArgs) Handles StopBtn.Click
        Vox.SpeakAsyncCancelAll()
        SpeechQueue.Clear()
        IsPaused = False
    End Sub

    Private Sub Vox_WordBoundary(sender As Object, e As SpeakProgressEventArgs) Handles Vox.SpeakProgress
        Dim spokenWord As String = e.Text
        Dim relativeStartIndex As Integer = e.CharacterPosition
        Dim absoluteStartIndex As Integer = currentReviewTextIndex + relativeStartIndex

        Debug.WriteLine($"[Vox_WordBoundary] Spoken Word: '{spokenWord}'")
        Debug.WriteLine($"[Vox_WordBoundary] Relative Start Index: {relativeStartIndex}")
        Debug.WriteLine($"[Vox_WordBoundary] Absolute Start Index: {absoluteStartIndex}")

        If Not isHighlighting Then
            isHighlighting = True
            HighlightWord(spokenWord, absoluteStartIndex)
        End If
    End Sub

    Private Sub HighlightWord(word As String, startIndex As Integer)
        Debug.WriteLine($"[HighlightWord] Word to Highlight: '{word}'")
        Debug.WriteLine($"[HighlightWord] Start Index: {startIndex}")

        If startIndex >= 0 AndAlso startIndex < ReviewTxt.Text.Length Then
            If startIndex + word.Length <= ReviewTxt.Text.Length Then
                Dim textToCompare As String = ReviewTxt.Text.Substring(startIndex, word.Length)
                Debug.WriteLine($"[HighlightWord] Text to Compare: '{textToCompare}'")
                If textToCompare.Equals(word, StringComparison.CurrentCultureIgnoreCase) Then
                    ReviewTxt.Select(startIndex, word.Length)
                    Application.DoEvents()
                    ReviewTxt.Invalidate()
                    ReviewTxt.SelectionColor = Color.Green

                    ' Restore after 200ms (adjust as needed)
                    Dim t As New Timer With {
                        .Interval = 100
                    }
                    AddHandler t.Tick, Sub()
                                           ReviewTxt.Select(startIndex, word.Length)
                                           ReviewTxt.SelectionColor = Color.Black
                                           t.Stop()
                                           t.Dispose()
                                           isHighlighting = False
                                       End Sub
                    t.Start()
                Else
                    Debug.WriteLine($"[HighlightWord] Comparison failed: '{textToCompare}' != '{word}'")
                    isHighlighting = False
                End If
            Else
                Debug.WriteLine($"[HighlightWord] End index out of bounds.")
                isHighlighting = False
            End If
        Else
            Debug.WriteLine($"[HighlightWord] Start index out of bounds.")
            isHighlighting = False
        End If
    End Sub

    Private Function GetVoice(voiceTag As String) As SpeechSynthesizer
        If voiceCache.ContainsKey(voiceTag) Then
            Return voiceCache(voiceTag) ' Return from cache
        Else
            ' Create a new voice object and add it to the cache
            Dim voice As New SpeechSynthesizer()
            Select Case voiceTag.ToLower()
                Case "narrator"
                    ' Get the VoiceInfo object for the narrator voice
                    Dim narratorVoiceInfo As VoiceInfo = Vox.GetInstalledVoices().FirstOrDefault(Function(v) v.VoiceInfo.Name = NarratorVoice)?.VoiceInfo
                    If narratorVoiceInfo IsNot Nothing Then
                        voice.SelectVoice(narratorVoiceInfo.Name)
                    Else
                        ' Handle the case where the narrator voice is not found
                        voice.SelectVoice(DefaultCharacterVoice) ' Or any other fallback voice
                    End If
                Case Else
                    If CharacterVoiceAssignments.ContainsKey(voiceTag) Then
                        ' Get the VoiceInfo object for the character voice
                        Dim characterVoiceInfo As VoiceInfo = Vox.GetInstalledVoices().FirstOrDefault(Function(v) v.VoiceInfo.Name = CharacterVoiceAssignments(voiceTag))?.VoiceInfo
                        If characterVoiceInfo IsNot Nothing Then
                            voice.SelectVoice(characterVoiceInfo.Name)
                        Else
                            ' Handle the case where the character voice is not found
                            voice.SelectVoice(DefaultCharacterVoice) ' Or any other fallback voice
                        End If
                    Else
                        ' Handle the case where the character voice is not found
                        voice.SelectVoice(DefaultCharacterVoice) ' Or any other fallback voice
                    End If
            End Select
            voiceCache.Add(voiceTag, voice)
            Return voice
        End If
    End Function

    Private Sub VoicesCBO_SelectedIndexChanged(sender As Object, e As EventArgs) Handles VoicesCBO.SelectedIndexChanged
        Try
            If VoicesCBO.SelectedItem IsNot Nothing Then
                NarratorVoice = VoicesCBO.SelectedItem.ToString()
                Vox.SelectVoice(NarratorVoice)
                'Vox.SpeakAsync("Hello, this is the narrator speaking.")
            End If
        Catch ex As Exception
            MessageBox.Show($"Error selecting voice: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadStcFile(filePath As String)
        Try

            Dim doc As New XmlDocument()
            doc.Load(filePath)

            ' Load RichTextBox content (unescape special characters)
            ' Load NarratorVoice
            Dim narratorNode As XmlNode = doc.SelectSingleNode("/StcFile/NarratorVoice")
            If narratorNode IsNot Nothing Then
                NarratorVoice = narratorNode.InnerText
            End If
            Try
                Dim rtfNode As XmlNode = doc.SelectSingleNode("/StcFile/RtfContent")
                If rtfNode IsNot Nothing Then
                    Dim escapedRtfContent As String = rtfNode.InnerText
                    Dim unescapedRtfContent As String = HttpUtility.HtmlDecode(escapedRtfContent) ' Unescape RTF content
                    Using ms As New MemoryStream(System.Text.Encoding.UTF8.GetBytes(unescapedRtfContent))
                        ReviewTxt.LoadFile(ms, RichTextBoxStreamType.RichText)
                    End Using
                End If
            Catch ex As Exception
                MessageBox.Show($"Error loading RTF: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

            ' Load voice assignments
            Try
                Dim assignmentsNode As XmlNode = doc.SelectSingleNode("/StcFile/VoiceAssignments")
                If assignmentsNode IsNot Nothing Then
                    CharacterVoiceAssignments.Clear() ' Clear existing assignments
                    For Each assignmentNode As XmlNode In assignmentsNode.SelectNodes("Assignment")
                        Dim character As String = assignmentNode.Attributes("Character").Value
                        Dim voice As String = assignmentNode.Attributes("Voice").Value
                        CharacterVoiceAssignments(character) = voice
                    Next
                End If
            Catch ex As Exception
                MessageBox.Show($"Error loading voice assignments: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try

        Catch ex As Exception
            MessageBox.Show($"Error loading .stc file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Form1.Text = "Story to Script Converter - " & IO.Path.GetFileName(filePath)
    End Sub

    Public Sub SaveStcFile(filePath As String)
        Try
            Dim doc As New XmlDocument()
            doc.AppendChild(doc.CreateElement("StcFile"))

            ' Save RichTextBox content (escape special characters)
            Dim rtfNode As XmlElement = doc.CreateElement("RtfContent")
            Using ms As New MemoryStream()
                ReviewTxt.SaveFile(ms, RichTextBoxStreamType.RichText)
                Dim rtfContent As String = System.Text.Encoding.UTF8.GetString(ms.ToArray())
                rtfNode.InnerText = System.Security.SecurityElement.Escape(rtfContent) ' Escape RTF content
            End Using
            doc.DocumentElement.AppendChild(rtfNode)

            ' Save voice assignments
            Dim assignmentsNode As XmlElement = doc.CreateElement("VoiceAssignments")
            For Each kvp As KeyValuePair(Of String, String) In CharacterVoiceAssignments
                Dim assignmentNode As XmlElement = doc.CreateElement("Assignment")
                assignmentNode.SetAttribute("Character", kvp.Key)
                assignmentNode.SetAttribute("Voice", kvp.Value)
                assignmentsNode.AppendChild(assignmentNode)
            Next
            doc.DocumentElement.AppendChild(assignmentsNode)
            ' Save NarratorVoice
            Dim narratorNode As XmlElement = doc.CreateElement("NarratorVoice")
            narratorNode.InnerText = NarratorVoice
            doc.DocumentElement.AppendChild(narratorNode)
            'doc.Save(filePath)
            doc.Save(New XmlTextWriter(filePath, System.Text.Encoding.UTF8))
        Catch ex As Exception
            MessageBox.Show($"Error saving .stc file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If MyXMLFile <> "" Then
            SaveStcFile(MyXMLFile)
        Else
            SaveAsToolStripMenuItem.PerformClick()
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim SFD As New SaveFileDialog With {
            .Filter = "STC Files (*.stc)|*.stc"
        }
        If SFD.ShowDialog() = DialogResult.OK Then
            MyXMLFile = SFD.FileName
            SaveStcFile(MyXMLFile)
        End If
    End Sub

End Class
