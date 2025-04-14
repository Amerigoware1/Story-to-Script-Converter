Imports System.IO
Imports System.Speech.Synthesis
Imports System.Text.RegularExpressions
Imports FuzzySharp

Public Class DialogueAssignmentControl '(WizardSteps(2)
    Private WithEvents Vox As New SpeechSynthesizer
    Public Property CharacterNames As List(Of String) ' To receive character names
    Private TextToAnalyze As String
    Private WithEvents CharacterComboBox As New ComboBox
    Private CharacterDialogueAssignments As New Dictionary(Of String, List(Of String))
    Private narratorLines As New List(Of String) ' New dictionary for narrator lines
    'Private currentFullContext As String = ""
    'Private currentRecentContext As String = ""
    Private lastAssignedCharacter As String = ""

    Private Sub DialogueAssignmentControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DialogueListView.Columns.Add("Voice", 100)
        CharacterComboBox.RightToLeft = RightToLeft.No
        CharacterComboBox.Width = 250
        Dim fileInputControl As FileInputControl = TryCast(Form1.WizardSteps(0), FileInputControl)
        If fileInputControl IsNot Nothing Then
            TextToAnalyze = fileInputControl.RTB1.Text
        End If

        For Each dialogueTuple As Tuple(Of String, Integer) In GetDialogueLines()
            Dim lineItem As New ListViewItem(dialogueTuple.Item1)
            lineItem.SubItems.Add("        ")
            lineItem.Tag = dialogueTuple.Item2 ' Store the start position in the Tag property
            DialogueListView.Items.Add(lineItem)
        Next
        Me.Controls.Add(CharacterComboBox)
    End Sub

    Public Sub PopulateCharacterComboBox() ' New subroutine to populate the ComboBox
        CharacterComboBox.Items.Clear() ' Clear any existing items
        If CharacterNames IsNot Nothing Then
            CharacterComboBox.Items.AddRange(CharacterNames.ToArray())
            CharacterComboBox.Items.Add("Unknown") ' Add an "Unknown" option
        End If
    End Sub

    'Private Sub DialogueListView_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DialogueListView.SelectedIndexChanged
    '    If DialogueListView.SelectedItems.Count > 0 Then
    '        Dim selectedItem As ListViewItem = DialogueListView.SelectedItems(0)
    '        Dim dialogueLine As String = selectedItem.SubItems(0).Text.Trim()
    '        Dim dialogueStartPos As Integer = CInt(selectedItem.Tag)

    '        ' Get the column header for the dialogue column
    '        Dim dialogueColumnHeader As ColumnHeader = DialogueListView.Columns(0)

    '        ' Calculate the ComboBox position (to the right of the dialogue column)
    '        Dim columnBounds As Rectangle = selectedItem.SubItems(0).Bounds
    '        Dim comboBoxX As Integer = columnBounds.Left + DialogueListView.Left + dialogueColumnHeader.Width
    '        Dim comboBoxY As Integer = columnBounds.Top + DialogueListView.Top

    '        ' Position and show the ComboBox
    '        CharacterComboBox.Bounds = New Rectangle(comboBoxX, comboBoxY, 150, columnBounds.Height)
    '        CharacterComboBox.Visible = True
    '        CharacterComboBox.SelectedIndex = -1 ' Clear the selection
    '        CharacterComboBox.BringToFront()

    '        ' Debugging output
    '        Debug.WriteLine($"ComboBox Bounds: {CharacterComboBox.Bounds}")

    '        ' Extract context outside quotes
    '        Dim context As String = ExtractContextOutsideQuotes(dialogueStartPos, dialogueLine.Length)

    '        ' Update full context and recent context
    '        currentFullContext &= context & " "
    '        currentRecentContext = context

    '        ' Auto-select character name
    '        AutoSelectCharacterName(dialogueLine, currentRecentContext)

    '    Else
    '        CharacterComboBox.Visible = False ' Hide the ComboBox if no item is selected
    '    End If
    'End Sub
    Private Sub DialogueListView_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DialogueListView.SelectedIndexChanged
        If DialogueListView.SelectedItems.Count > 0 Then
            Dim selectedItem As ListViewItem = DialogueListView.SelectedItems(0)
            Dim dialogueLine As String = selectedItem.SubItems(0).Text ' Keep original quotes for potential lookups if needed
            Dim dialogueStartPos As Integer = CInt(selectedItem.Tag)
            'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Dialogue Line: {dialogueLine}" & vbNewLine)
            ' Get the column header for the dialogue column
            Dim dialogueColumnHeader As ColumnHeader = DialogueListView.Columns(0)

            ' Calculate the ComboBox position (to the right of the dialogue column)
            Dim columnBounds As Rectangle = selectedItem.SubItems(0).Bounds ' Use bounds of the first subitem
            Dim comboBoxX As Integer = columnBounds.Left + DialogueListView.Left + dialogueColumnHeader.Width + 5 ' Added small offset
            Dim comboBoxY As Integer = columnBounds.Top + DialogueListView.Top

            ' Position and show the ComboBox
            CharacterComboBox.Bounds = New Rectangle(comboBoxX, comboBoxY, 150, columnBounds.Height)
            CharacterComboBox.Visible = True
            CharacterComboBox.SelectedIndex = -1 ' Clear the selection initially
            CharacterComboBox.BringToFront()

            ' --- Extract Context More Specifically ---
            ' Context immediately before and after the quote (within ~1 sentence boundary)
            Dim adjacentContext As String = ExtractContextOutsideQuotes(dialogueStartPos, dialogueLine.Length) '[cite: 8, 9]
            'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Adjacent Context: {adjacentContext}" & vbNewLine)
            ' Optional: Wider context (e.g., fixed number of words/chars if ExtractContext... is too narrow)
            Dim charsBefore As Integer = 150 ' Adjust as needed
            Dim charsAfter As Integer = 100  ' Adjust as needed
            Dim widerContextStart As Integer = Math.Max(0, dialogueStartPos - charsBefore)
            Dim widerContextLengthBefore As Integer = dialogueStartPos - widerContextStart
            Dim widerContextBefore As String = TextToAnalyze.Substring(widerContextStart, widerContextLengthBefore)
            'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Wider Context Before: {widerContextBefore}" & vbNewLine)
            Dim widerContextAfterStart As Integer = dialogueStartPos + dialogueLine.Length
            Dim widerContextLengthAfter As Integer = Math.Min(charsAfter, TextToAnalyze.Length - widerContextAfterStart)
            Dim widerContextAfter As String = ""
            If widerContextLengthAfter > 0 Then
                widerContextAfter = TextToAnalyze.Substring(widerContextAfterStart, widerContextLengthAfter)
                'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Wider Context After: {widerContextAfter}" & vbNewLine)
            End If
            ' Clean wider context from quotes
            widerContextBefore = Regex.Replace(widerContextBefore, """[^""]+""", "")
            'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Wider Context Before Cleaned: {widerContextBefore}" & vbNewLine)
            widerContextAfter = Regex.Replace(widerContextAfter, """[^""]+""", "")
            'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Wider Context After Cleaned: {widerContextAfter}" & vbNewLine)
            Dim widerContext As String = (widerContextBefore & " " & widerContextAfter).Trim()
            'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Wider Context: {widerContext}" & vbNewLine)

            ' --- Auto-select character name using the new context and persistence ---
            ' Pass both adjacent and wider context, plus position info and last speaker
            AutoSelectCharacterName(dialogueLine, adjacentContext, widerContext, dialogueStartPos, dialogueLine.Length, lastAssignedCharacter)
            'IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: === AutoSelectCharacterName End ===" & vbNewLine)
        Else
            CharacterComboBox.Visible = False ' Hide the ComboBox if no item is selected
        End If
    End Sub
    Private Function ExtractContextOutsideQuotes(dialogueStartPos As Integer, dialogueLength As Integer) As String
        Dim context As String
        Dim beforeContext As String
        Dim afterContext As String

        ' Extract context before the dialogue
        Dim beforeEndPos As Integer = dialogueStartPos
        Dim beforeStartPos As Integer = FindPreviousSentenceEnd(TextToAnalyze, beforeEndPos)
        beforeContext = TextToAnalyze.Substring(beforeStartPos, beforeEndPos - beforeStartPos)

        ' Extract context after the dialogue
        Dim afterStartPos As Integer = dialogueStartPos + dialogueLength
        Dim afterEndPos As Integer = FindNextSentenceEnd(TextToAnalyze, afterStartPos)
        afterContext = TextToAnalyze.Substring(afterStartPos, afterEndPos - afterStartPos)

        ' Remove quoted text from context
        beforeContext = Regex.Replace(beforeContext, """[^""]+""", "")
        afterContext = Regex.Replace(afterContext, """[^""]+""", "")

        ' Combine and trim context (adjust as needed)
        context = beforeContext.Trim() & " " & afterContext.Trim()
        Return context.Trim()
    End Function

    Private Function FindPreviousSentenceEnd(text As String, endPos As Integer) As Integer
        Dim sentenceEndings As String = ".!?"
        For i As Integer = endPos - 1 To 0 Step -1
            If sentenceEndings.Contains(text(i)) Then
                Return i + 1 ' Return the position after the sentence ending
            End If
        Next
        Return 0 ' Beginning of text
    End Function

    Private Function FindNextSentenceEnd(text As String, startPos As Integer) As Integer
        Dim sentenceEndings As String = ".!?"
        For i As Integer = startPos To text.Length - 1
            If sentenceEndings.Contains(text(i)) Then
                Return i + 1 ' Return the position after the sentence ending
            End If
        Next
        Return text.Length ' End of text
    End Function

    Private Sub DialogueListView_Leave(sender As Object, e As EventArgs) Handles DialogueListView.Leave
        If Not CharacterComboBox.Focused Then
            CharacterComboBox.Visible = False ' Hide the ComboBox when the ListView loses focus
        End If
    End Sub

    Private Function GetDialogueLines() As List(Of Tuple(Of String, Integer)) ' Return a list of tuples (dialogue, start position)
        Dim lines As New List(Of Tuple(Of String, Integer))()
        Dim pattern As String = """[^""]+""" ' Matches text within double quotes
        Dim matches As MatchCollection = Regex.Matches(TextToAnalyze, pattern)
        For Each match As Match In matches
            lines.Add(New Tuple(Of String, Integer)(match.Value, match.Index)) ' Store dialogue and its start position
        Next
        Return lines
    End Function

    'Private Sub CharacterComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CharacterComboBox.SelectedIndexChanged
    '    If DialogueListView.SelectedItems.Count > 0 AndAlso CharacterComboBox.SelectedItem IsNot Nothing Then
    '        Dim selectedItem As ListViewItem = DialogueListView.SelectedItems(0)
    '        'Dim selectedCharacter As String = CharacterComboBox.SelectedItem.ToString()
    '        'Dim selectedVoice As String = CharacterComboBox.SelectedItem.ToString()
    '        Dim selectedCharacter As String = CharacterComboBox.SelectedItem.ToString()
    '        Dim selectedVoice As String = selectedCharacter ' Assuming voice is the same as character name for now
    '        ' Populate narratorLines BEFORE populating CharacterDialogueAssignments
    '        Dim allLines As New List(Of String)(Regex.Split(TextToAnalyze, "\r?\n")) ' Splits on \r\n or \n
    '        Debug.WriteLine($"allLines Count: {allLines.Count}") ' Check allLines count
    '        My.Computer.FileSystem.WriteAllText(Form1.Log, $"allLines Count: {allLines.Count}" & vbCrLf, True)
    '        narratorLines.Clear() ' Clear the list before repopulating
    '        For Each line As String In allLines
    '            ' Check for exact match with normalized strings
    '            If Not CharacterDialogueAssignments.Any(Function(kvp) kvp.Value.Any(Function(dialogue) line.Trim().ToLower() = dialogue.Trim().ToLower())) Then
    '                narratorLines.Add(line)
    '                Debug.WriteLine($"Narrator Line: {line}") ' Check narrator lines
    '            End If
    '        Next
    '        ' Store the entire line, including leading/trailing text
    '        If Not CharacterDialogueAssignments.ContainsKey(selectedCharacter) Then
    '            CharacterDialogueAssignments(selectedCharacter) = New List(Of String)
    '        End If
    '        CharacterDialogueAssignments(selectedCharacter).Add(selectedItem.Text)



    '        selectedItem.SubItems(1).Text = CharacterComboBox.SelectedItem.ToString()
    '    End If

    '    Form1.NextButton.Enabled = True
    'End Sub
    Private Sub CharacterComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CharacterComboBox.SelectedIndexChanged
        If DialogueListView.SelectedItems.Count > 0 AndAlso CharacterComboBox.SelectedItem IsNot Nothing Then
            Dim selectedItem As ListViewItem = DialogueListView.SelectedItems(0)
            Dim selectedCharacter As String = CharacterComboBox.SelectedItem.ToString()
            Dim dialogueLineText As String = selectedItem.SubItems(0).Text ' The dialogue line itself

            ' --- Update the ListView ---
            selectedItem.SubItems(1).Text = selectedCharacter

            ' --- Update Assignments (Simplified - Review your original logic here) ---
            ' Ensure this logic correctly adds the dialogue to the selected character
            ' and handles narrator lines appropriately based on your overall design.
            ' The example below assumes CharacterDialogueAssignments stores dialogue lines per character.

            ' Remove line from other characters if previously assigned
            For Each kvp As KeyValuePair(Of String, List(Of String)) In CharacterDialogueAssignments
                If kvp.Key <> selectedCharacter Then
                    kvp.Value.Remove(dialogueLineText)
                End If
            Next

            ' Add to the selected character's list
            If Not CharacterDialogueAssignments.ContainsKey(selectedCharacter) Then
                CharacterDialogueAssignments(selectedCharacter) = New List(Of String)()
            End If
            If Not CharacterDialogueAssignments(selectedCharacter).Contains(dialogueLineText) Then
                CharacterDialogueAssignments(selectedCharacter).Add(dialogueLineText)
            End If

            ' *** NEW: Update the last assigned character for persistence ***
            If Not selectedCharacter.Equals("Unknown", StringComparison.OrdinalIgnoreCase) Then
                lastAssignedCharacter = selectedCharacter
                'IO.File.AppendAllText("log.txt", $"{DateTime.Now}: User/Auto Selection Confirmed. Last Speaker set to: {lastAssignedCharacter}" & vbNewLine)
            End If

            ' --- Narrator Line Logic (Review Needed) ---
            ' Your original code repopulated narratorLines here[cite: 17, 18, 19].
            ' This seems inefficient. Consider updating narratorLines only when
            ' the entire assignment process is complete, or manage it differently.
            ' Repopulating it on every selection change might be slow and logically complex.
            ' For now, commenting out the original narrator logic from here:
            ' narratorLines.Clear()
            ' Dim allLines As New List(Of String)(Regex.Split(TextToAnalyze, "\r?\n"))
            ' For Each line As String In allLines
            '     If Not CharacterDialogueAssignments.Any(...) Then
            '         narratorLines.Add(line)
            '     End If
            ' Next

            ' --- Enable Next Button ---
            Form1.NextButton.Enabled = True
        End If
    End Sub
    Public Function GetNarratorLines() As List(Of String)
        Return narratorLines
    End Function

    Public Function GetCharacterDialogueAssignments() As Dictionary(Of String, List(Of String))
        Return CharacterDialogueAssignments
    End Function

    'Private Sub AutoSelectCharacterName(dialogueLine As String, fullContext As String)
    '    ' Load clue words from file
    '    Dim clueWords As New List(Of String)
    '    Try
    '        Dim clueFilePath As String = My.Application.Info.DirectoryPath & "\ClueWords.txt"
    '        If File.Exists(clueFilePath) Then
    '            clueWords.AddRange(File.ReadAllLines(clueFilePath))
    '        End If
    '    Catch ex As Exception
    '        IO.File.AppendAllText("log.txt", $"{DateTime.Now}: Error loading clue words: {ex.Message}" & vbNewLine)
    '    End Try

    '    ' Debug log start
    '    IO.File.AppendAllText("log.txt", $"{DateTime.Now}: AutoSelectCharacterName triggered." & vbNewLine)
    '    IO.File.AppendAllText("log.txt", $"Dialogue Line: {dialogueLine}" & vbNewLine)
    '    IO.File.AppendAllText("log.txt", $"Full Context: {fullContext}" & vbNewLine)

    '    Dim actorScores As New Dictionary(Of String, Integer)()

    '    ' --- Sentence Splitting and Context ---
    '    Dim contextSentences As List(Of String) = Regex.Split(fullContext, "(?<=[.!?])\s+").ToList()
    '    Dim recentContext As String = String.Join(" ", contextSentences.Skip(Math.Max(0, contextSentences.Count - 10))) ' Expanded context

    '    IO.File.AppendAllText("log.txt", $"Recent Context: {recentContext}" & vbNewLine)

    '    ' --- Match Scoring (Corrected) ---
    '    For Each actor As String In CharacterComboBox.Items
    '        Dim name As String = actor.ToString()
    '        actorScores(name) = 0

    '        ' Direct match scoring
    '        If recentContext.Contains(name) Then
    '            actorScores(name) += 50 ' Increased score for direct match
    '        End If

    '        ' Fuzzy matching scoring
    '        actorScores(name) += Fuzz.Ratio(name, recentContext) * 1.0 ' Increased fuzzy matching weight

    '        ' Clue word scoring
    '        For Each clueWord As String In clueWords
    '            Dim regexPattern As String = $"{name}.*?\b{clueWord}\b" ' Flexible clue word matching
    '            If Regex.IsMatch(recentContext, regexPattern, RegexOptions.IgnoreCase) Then
    '                actorScores(name) += 50 ' Very high score for clue word match
    '            End If
    '        Next
    '    Next

    '    ' --- Select Best Match ---
    '    Dim bestMatch As String = ""
    '    Dim bestMatchScore As Integer = -1

    '    For Each actor As String In actorScores.Keys
    '        If actorScores(actor) > bestMatchScore Then
    '            bestMatchScore = actorScores(actor)
    '            bestMatch = actor
    '        End If
    '        IO.File.AppendAllText("log.txt", $"Actor: {actor}, Score: {actorScores(actor)}" & vbNewLine)
    '    Next

    '    If bestMatch <> "" And bestMatchScore > 50 Then 'Adjust the threshold
    '        CharacterComboBox.SelectedItem = bestMatch
    '        IO.File.AppendAllText("log.txt", $"Best match: {bestMatch}, Score: {bestMatchScore}" & vbNewLine)
    '    Else
    '        IO.File.AppendAllText("log.txt", $"No confident match. Keeping default selection." & vbNewLine)
    '    End If
    'End Sub
    ' *** MODIFIED Signature to accept more specific context and persistence info ***
    Private Sub AutoSelectCharacterName(dialogueLine As String, adjacentContext As String, widerContext As String, ByVal dialogueStartPos As Integer, ByVal dialogueLength As Integer, ByVal previousSpeaker As String)

        ' Load clue words from file (keep this logic) [cite: 21, 22]
        Dim clueWords As New List(Of String)
        Try
            Dim clueFilePath As String = My.Application.Info.DirectoryPath & "\ClueWords.txt"
            If File.Exists(clueFilePath) Then
                clueWords.AddRange(File.ReadAllLines(clueFilePath))
            End If
        Catch ex As Exception
            IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: Error loading clue words: {ex.Message}" & vbNewLine)
        End Try

        ' Combine contexts for general checks, but prioritize adjacent context
        Dim combinedContext As String = (adjacentContext & " " & widerContext).Trim()

        ' Log inputs for debugging
        IO.File.AppendAllText(Form1.Log, $"{DateTime.Now}: === AutoSelectCharacterName Start ===" & vbNewLine)
        IO.File.AppendAllText(Form1.Log, $"Dialogue Line: {dialogueLine}" & vbNewLine)
        IO.File.AppendAllText(Form1.Log, $"Adjacent Context: {adjacentContext}" & vbNewLine)
        IO.File.AppendAllText(Form1.Log, $"Wider Context: {widerContext}" & vbNewLine)
        IO.File.AppendAllText(Form1.Log, $"Previous Speaker: {previousSpeaker}" & vbNewLine)


        Dim actorScores As New Dictionary(Of String, Integer)()
        Dim secondBestScore As Integer = -1 ' For relative threshold check

        ' --- Text immediately before/after quote ---
        Dim textImmediatelyBefore As String = ""
        Dim beforeCheckLength As Integer = 30 ' How many chars before quote to check for "Name said" patterns
        Dim beforeStart As Integer = Math.Max(0, dialogueStartPos - beforeCheckLength)
        If beforeStart < dialogueStartPos Then
            textImmediatelyBefore = TextToAnalyze.Substring(beforeStart, dialogueStartPos - beforeStart).Trim()
            textImmediatelyBefore = Regex.Replace(textImmediatelyBefore, """[^""]+""", "") ' Remove quotes
        End If

        Dim textImmediatelyAfter As String = ""
        Dim afterCheckLength As Integer = 30 ' How many chars after quote to check for "said Name" patterns
        Dim afterStart As Integer = dialogueStartPos + dialogueLength
        If afterStart < TextToAnalyze.Length Then
            Dim afterEnd As Integer = Math.Min(TextToAnalyze.Length, afterStart + afterCheckLength)
            textImmediatelyAfter = TextToAnalyze.Substring(afterStart, afterEnd - afterStart).Trim()
            textImmediatelyAfter = Regex.Replace(textImmediatelyAfter, """[^""]+""", "") ' Remove quotes
        End If

        IO.File.AppendAllText(Form1.Log, $"Text Immediately Before ({beforeCheckLength} chars): {textImmediatelyBefore}" & vbNewLine)
        IO.File.AppendAllText(Form1.Log, $"Text Immediately After ({afterCheckLength} chars): {textImmediatelyAfter}" & vbNewLine)


        ' --- Score Actors ---
        For Each actorObject As Object In CharacterComboBox.Items
            Dim name As String = actorObject.ToString()

            ' *** Skip "Unknown" character ***
            If name.Equals("Unknown", StringComparison.OrdinalIgnoreCase) Then
                Continue For
            End If

            actorScores(name) = 0 ' Initialize score

            ' --- Scoring Weights ---
            Const DIRECT_MATCH_SCORE As Integer = 60         ' High score for direct name mention in relevant context
            Const ADJACENT_CLUE_SCORE As Integer = 100       ' Highest score for patterns like "Name said," right next to quote
            Const CONTEXT_CLUE_SCORE As Integer = 40         ' Score for name + clue word in wider context
            Const PERSISTENCE_BONUS As Integer = 50          ' Bonus if this actor spoke last
            Const FUZZY_MATCH_WEIGHT As Double = 0.2         ' Reduced weight for general fuzzy matching

            ' 1. Speaker Persistence Bonus
            If Not String.IsNullOrEmpty(previousSpeaker) AndAlso name.Equals(previousSpeaker, StringComparison.OrdinalIgnoreCase) Then
                actorScores(name) += PERSISTENCE_BONUS
                IO.File.AppendAllText(Form1.Log, $"    Score: +{PERSISTENCE_BONUS} (Persistence Bonus for {name})" & vbNewLine)
            End If

            ' 2. Direct Name Match (prioritize adjacent context)
            If adjacentContext.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0 Then
                actorScores(name) += DIRECT_MATCH_SCORE
                IO.File.AppendAllText(Form1.Log, $"    Score: +{DIRECT_MATCH_SCORE} (Direct Match in Adjacent Context for {name})" & vbNewLine)
            ElseIf widerContext.IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0 Then
                actorScores(name) += CInt(DIRECT_MATCH_SCORE * 0.5) ' Lower score for match in wider context only
                IO.File.AppendAllText(Form1.Log, $"    Score: +{CInt(DIRECT_MATCH_SCORE * 0.5)} (Direct Match in Wider Context for {name})" & vbNewLine)
            End If


            ' 3. Clue Word Scoring (Prioritize Adjacent Patterns)
            Dim clueFound As Boolean = False
            For Each clueWord As String In clueWords
                ' Pattern 1: "{Name} [said]," immediately BEFORE quote
                Dim patternBefore As String = $"\b{Regex.Escape(name)}\b\s*,?\s*\b{Regex.Escape(clueWord)}\b\s*$" ' Ends with name+clue before quote
                If Regex.IsMatch(textImmediatelyBefore, patternBefore, RegexOptions.IgnoreCase) Then
                    actorScores(name) += ADJACENT_CLUE_SCORE
                    IO.File.AppendAllText(Form1.Log, $"    Score: +{ADJACENT_CLUE_SCORE} (Adjacent Clue BEFORE for {name} + {clueWord})" & vbNewLine)
                    clueFound = True
                    Exit For ' Strongest match, no need to check others for this actor
                End If

                ' Pattern 2: ", [said] {Name}" immediately AFTER quote
                Dim patternAfter As String = $"^\s*,?\s*\b{Regex.Escape(clueWord)}\b\s+\b{Regex.Escape(name)}\b" ' Starts with clue+name after quote
                If Regex.IsMatch(textImmediatelyAfter, patternAfter, RegexOptions.IgnoreCase) Then
                    actorScores(name) += ADJACENT_CLUE_SCORE
                    IO.File.AppendAllText(Form1.Log, $"    Score: +{ADJACENT_CLUE_SCORE} (Adjacent Clue AFTER for {name} + {clueWord})" & vbNewLine)
                    clueFound = True
                    Exit For ' Strongest match
                End If
            Next

            ' If no strong adjacent clue found, check wider context (less reliable)
            If Not clueFound Then
                For Each clueWord As String In clueWords
                    ' Original less specific pattern, applied to combined context, lower score
                    Dim regexPatternContext As String = $"\b{Regex.Escape(name)}\b.*?(\b{Regex.Escape(clueWord)}\b)"
                    If Regex.IsMatch(combinedContext, regexPatternContext, RegexOptions.IgnoreCase) Then
                        actorScores(name) += CONTEXT_CLUE_SCORE
                        IO.File.AppendAllText(Form1.Log, $"    Score: +{CONTEXT_CLUE_SCORE} (Context Clue for {name} + {clueWord})" & vbNewLine)
                        ' Don't Exit For here, allow multiple weaker clues perhaps? Or Exit For if one is enough. Decide based on testing.
                    End If
                    ' Optional: Add pattern for "\bsaid\b.*?{Name}" in context too?
                Next
            End If


            ' 4. Fuzzy Matching (Reduced Influence)
            Dim fuzzyScore As Integer = CInt(Fuzz.Ratio(name, adjacentContext) * FUZZY_MATCH_WEIGHT) ' Apply fuzzy only to adjacent context
            actorScores(name) += fuzzyScore
            If fuzzyScore > 0 Then IO.File.AppendAllText(Form1.Log, $"    Score: +{fuzzyScore} (Fuzzy Match Score for {name})" & vbNewLine)

        Next ' Next actor

        ' --- Select Best Match ---
        Dim bestMatch As String = ""
        Dim bestMatchScore As Integer = 0 ' Use 0 as base, not -1

        ' Find best and second best scores
        Dim sortedScores = actorScores.OrderByDescending(Function(kvp) kvp.Value).ToList()

        If sortedScores.Count > 0 Then
            bestMatch = sortedScores(0).Key
            bestMatchScore = sortedScores(0).Value
            IO.File.AppendAllText(Form1.Log, $"Top Candidate: {bestMatch}, Score: {bestMatchScore}" & vbNewLine)

            If sortedScores.Count > 1 Then
                secondBestScore = sortedScores(1).Value
                IO.File.AppendAllText(Form1.Log, $"Second Candidate: {sortedScores(1).Key}, Score: {secondBestScore}" & vbNewLine)
            Else
                secondBestScore = -1 ' No second candidate
            End If
        End If

        ' --- Confidence Check ---
        Dim confidenceThreshold As Double = 1.5 ' Best score must be 1.5x the second best
        Dim absoluteMinimumScore As Integer = 40 ' Absolute minimum score required, even if confidence is high

        Dim confidentMatch As Boolean = False
        If bestMatchScore >= absoluteMinimumScore Then
            If secondBestScore < 0 Then ' Only one candidate with score > 0
                confidentMatch = True
            ElseIf secondBestScore = 0 AndAlso bestMatchScore > 0 Then ' Handle case where second best is 0
                confidentMatch = True
            ElseIf secondBestScore > 0 AndAlso bestMatchScore >= secondBestScore * confidenceThreshold Then ' Relative check
                confidentMatch = True
            End If
        End If


        If confidentMatch Then
            CharacterComboBox.SelectedItem = bestMatch
            ' *** IMPORTANT: Update lastAssignedCharacter state AFTER selection ***
            ' This happens in CharacterComboBox_SelectedIndexChanged now when user confirms or changes.
            IO.File.AppendAllText(Form1.Log, $"==> Confident Match Selected: {bestMatch}, Score: {bestMatchScore} (Threshold Met)" & vbNewLine)
        Else
            ' Keep default selection (usually blank)
            IO.File.AppendAllText(Form1.Log, $"==> No Confident Match. Score: {bestMatchScore}. (Thresholds: Abs>{absoluteMinimumScore}, Rel>{confidenceThreshold}x)" & vbNewLine)
            ' Optional: Select "Unknown" explicitly if no match, or leave blank
            ' CharacterComboBox.SelectedItem = "Unknown" ' Or find "Unknown" item if it exists
        End If
        IO.File.AppendAllText(Form1.Log, $"=== AutoSelectCharacterName End ===" & vbNewLine & vbNewLine)

    End Sub

    ' --- Ensure ExtractContextOutsideQuotes and related functions are present ---
    ' [cite: 9, 10, 11, 12, 13, 14]
    ' Make sure these functions exist as shown in the original file content.
    'Private Function ExtractContextOutsideQuotes(dialogueStartPos As Integer, dialogueLength As Integer) As String
    '    ' ... (Implementation as provided in the file)...
    'End Function

    'Private Function FindPreviousSentenceEnd(text As String, endPos As Integer) As Integer
    '    ' ... (Implementation as provided in the file)[cite: 11, 12]...
    'End Function

    'Private Function FindNextSentenceEnd(text As String, startPos As Integer) As Integer
    '    ' ... (Implementation as provided in the file)[cite: 12, 13, 14]...
    'End Function
End Class
