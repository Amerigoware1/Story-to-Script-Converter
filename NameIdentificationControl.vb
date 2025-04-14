Imports System.IO, System.Text.RegularExpressions
Imports System.Diagnostics
Imports System.Diagnostics.Eventing.Reader

Public Class NameIdentificationControl '(WizardSteps(1)

    Private NameDict As New List(Of String)
    Private CommonWords As New List(Of String)
    Private Titles As New List(Of String)
    Private ClueWords As New List(Of String)
    Public TextToAnalyze As String

    Private Sub NameIdentificationControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        StatusLbl.Parent = ProgressBar1
        StatusLbl.BackColor = Color.Transparent
        StatusLbl.Location = New Point(0, 0)
        Form1.NextButton.Enabled = False
        Try
            If File.Exists("Name List.txt") Then
                Dim names As String() = File.ReadAllText("Name List.txt").Split({" ", vbCrLf}, StringSplitOptions.RemoveEmptyEntries)
                ProgressBar1.Maximum = names.Length
                StatusLbl.Text = "Loading Name List..."
                For Each name As String In names
                    NameDict.Add(name)
                    ProgressBar1.Value += 1
                Next
            End If

            If File.Exists("en-US.txt") Then
                Dim Words As String() = File.ReadAllLines("en-US.txt")
                ProgressBar1.Value = 0
                ProgressBar1.Maximum = Words.Length
                StatusLbl.Text = "Loading Common Words..."
                For Each word As String In Words
                    CommonWords.Add(word)
                    ProgressBar1.Value += 1
                Next
            End If

            If File.Exists("Titles.txt") Then
                Dim Honorifics As String() = File.ReadAllLines("Titles.txt")
                ProgressBar1.Value = 0
                ProgressBar1.Maximum = Honorifics.Length
                StatusLbl.Text = "Loading Titles..."
                For Each t As String In Honorifics
                    Titles.Add(t)
                    ProgressBar1.Value += 1
                Next
            End If

            If File.Exists("ClueWords.txt") Then
                Dim Clues As String() = File.ReadAllLines("ClueWords.txt")
                ProgressBar1.Value = 0
                ProgressBar1.Maximum = Clues.Length
                StatusLbl.Text = "Loading Clue Words..."
                For Each c As String In Clues
                    ClueWords.Add(c)
                    ProgressBar1.Value += 1
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading Name List: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        ProgressBar1.Value = 0
        StatusLbl.Text = ""
        TextToAnalyze = NormalizeLineEndings(TextToAnalyze)
        LoadPotentialNames()
    End Sub
    ' Helper function to normalize line endings to CRLF (Windows style)
    Private Function NormalizeLineEndings(text As String) As String
        Return Regex.Replace(text, "(\r\n|\n|\r)", Environment.NewLine)
    End Function

    Private Sub LoadPotentialNames()
        Try
            If Not String.IsNullOrEmpty(TextToAnalyze) Then
                Dim potentialNames As List(Of String) = ExtractNames(TextToAnalyze, NameDict) ' Pass NameDict
                'Dim potentialNames As List(Of String) = ExtractSpeakingCharacterNames(TextToAnalyze, NameDict) ' Pass NameDict
                PotentialNameListBox.Items.Clear()
                For Each name As String In potentialNames
                    PotentialNameListBox.Items.Add(name)
                Next
            End If
        Catch ex As Exception
            MessageBox.Show("Error extracting names: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function CleanName(name As String) As String
        ' Keep apostrophes, hyphens, and periods (e.g., in honorifics), but remove other non-alphanumeric characters.
        ' Preserve Unicode letters and numbers.
        Dim Str As String
        If Titles.Contains(name, StringComparer.OrdinalIgnoreCase) Then
            Str = name
        Else
            ' Remove any character that is not a Unicode letter, digit, hyphen, apostrophe, period, or space
            Str = Regex.Replace(name, "[^\p{L}\p{N}\-'. ]+", "")
        End If

        ' Normalize line breaks and other whitespace characters to a single space
        Str = Regex.Replace(Str, "[\r\n]+", " ")

        ' Collapse multiple spaces into one
        Str = Regex.Replace(Str, "\s+", " ")

        ' Trim spaces at the ends
        Return Str.Trim()
    End Function


    Private Function ExtractSpeakingCharacterNames(text As String, nameDict As List(Of String)) As List(Of String)
        Dim names As New List(Of String)()
        ' Regex to find text within double quotes (dialogue)
        Dim dialogueRegex As New Regex("""([^""]*)""")
        ' Regex to find capitalized words (potential names)
        Dim potentialNameRegex As New Regex("\b[A-Z][\w\-\']*\b")
        ' Regex to find sequences of one or more capitalized words
        Dim potentialMultiWordNameRegex As New Regex("\b[A-Z][\w\-\']+(?:\s+[A-Z][\w\-\']+)*\b")

        ' Extract dialogue portions
        Dim dialogueMatches As MatchCollection = dialogueRegex.Matches(text)
        ProgressBar1.Maximum = dialogueMatches.Count
        StatusLbl.Text = "Extracting names from dialogue..."
        For Each dialogueMatch As Match In dialogueMatches
            Dim dialogue As String = dialogueMatch.Value.Trim("""") ' Get the text inside the quotes

            ' Find potential names within the dialogue
            Dim potentialNameMatches As MatchCollection = potentialNameRegex.Matches(dialogue)

            For Each potentialNameMatch As Match In potentialNameMatches
                Dim potentialName As String = potentialNameMatch.Value
                Dim cleanedName As String = CleanName(potentialName)

                ' Check if it's in the name dictionary and not a common word or title
                If Not Titles.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
                   nameDict.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
                   Not CommonWords.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
                   Not names.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) Then
                    names.Add(cleanedName)
                    ProgressBar1.Value += 1
                    ProgressBar1.Update()
                    File.AppendAllText(Form1.Log, $"Added potential name from dialogue: {cleanedName}{vbNewLine}")
                    Application.DoEvents()
                End If
            Next
        Next
        ProgressBar1.Value = 0

        ' Find names *outside* dialogue, using clue words
        Dim textWithoutDialogue As String = dialogueRegex.Replace(text, "")
        For Each clueWord As String In ClueWords
            Dim clueWordRegex As New Regex("\b" & clueWord & "\s+([A-Z][\w\-\']+)") 'Finds capitalized word after a clue word
            Dim clueWordMatches As MatchCollection = clueWordRegex.Matches(textWithoutDialogue)
            ProgressBar1.Maximum += clueWordMatches.Count
            StatusLbl.Text = "Extracting names from clue words..."
            For Each clueWordMatch As Match In clueWordMatches
                Dim potentialName As String = clueWordMatch.Groups(1).Value 'Get the captured group (the name)
                Dim cleanedName As String = CleanName(potentialName)

                ' Check if it's in the name dictionary and not a common word or title
                If Not Titles.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
                   nameDict.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
                   Not CommonWords.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
                   Not names.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) Then
                    names.Add(cleanedName)
                    ProgressBar1.Value += 1
                    ProgressBar1.Update()
                    File.AppendAllText(Form1.Log, $"Added potential name from clue word: {cleanedName}{vbNewLine}")
                    Application.DoEvents()
                End If
            Next
            ProgressBar1.Value = 0
            'Check for multi-word names after clue words
            Dim multiClueWordRegex As New Regex("\b" & clueWord & "\s+([A-Z][\w\-\']+(?:\s+[A-Z][\w\-\']+)*)")
            Dim multiClueWordMatches As MatchCollection = multiClueWordRegex.Matches(textWithoutDialogue)
            ProgressBar1.Maximum += multiClueWordMatches.Count
            StatusLbl.Text = "Extracting multi-word names from clue words..."
            For Each multiClueWordMatch As Match In multiClueWordMatches
                Dim potentialMultiWordName As String = multiClueWordMatch.Groups(1).Value
                Dim wordsInName As String() = potentialMultiWordName.Split(New Char() {" "}, StringSplitOptions.RemoveEmptyEntries)
                Dim allNonTitleWordsInDict As Boolean = False
                Dim hasNonTitleWord As Boolean = False

                If wordsInName.Length > 1 Then
                    For Each word As String In wordsInName
                        Dim cleanedWord As String = CleanName(word)
                        If Not Titles.Contains(cleanedWord, StringComparer.OrdinalIgnoreCase) Then
                            hasNonTitleWord = True
                            If Not nameDict.Contains(cleanedWord, StringComparer.OrdinalIgnoreCase) Then
                                allNonTitleWordsInDict = False
                                Exit For
                            Else
                                allNonTitleWordsInDict = True
                            End If
                        End If
                    Next
                    If hasNonTitleWord AndAlso allNonTitleWordsInDict AndAlso Not names.Contains(potentialMultiWordName, StringComparer.OrdinalIgnoreCase) Then
                        names.Add(potentialMultiWordName)
                        ProgressBar1.Value += 1
                        ProgressBar1.Update()
                        File.AppendAllText(Form1.Log, $"Added potential multi-word name: {potentialMultiWordName}{vbNewLine}")
                        Application.DoEvents()

                    End If
                End If
            Next
        Next
        StatusLbl.Text = ""
        ProgressBar1.Value = 0
        Return names
    End Function

    Private Function ExtractNames(text As String, nameDict As List(Of String)) As List(Of String)

        Dim names As New List(Of String)()
        ' Regex to find text within double quotes
        Dim dialogueRegex As New Regex("""([^""]*)""")
        ' Regex to find single words starting with a capital letter
        Dim singleWordNameRegex As New Regex("\b[A-Z][\w\-\']*\b")
        ' Regex to find sequences of one or more capitalized words
        Dim potentialMultiWordNameRegex As New Regex("\b[A-Z][\w\-\']+(?:\s+[A-Z][\w\-\']+)*\b")

        ' Remove text within double quotes to only process text outside dialogue
        Dim textWithoutDialogue As String = dialogueRegex.Replace(text, "")

        StatusLbl.Text = "Extracting names from story..."

        ' Extract and process single-word names
        Dim singleWordMatches As MatchCollection = singleWordNameRegex.Matches(textWithoutDialogue)
        ProgressBar1.Maximum = singleWordMatches.Count
        For Each match As Match In singleWordMatches
            Dim potentialName As String = match.Value
            Dim cleanedName As String = CleanName(potentialName)

            ' Incorporate honorifics check for single-word names (remains the same)
            If Not Titles.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) Then
                If nameDict.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso Not CommonWords.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
                   Not names.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) Then
                    names.Add(cleanedName)
                    IO.File.AppendAllText(Form1.Log, $"Added potential name: {cleanedName}{vbNewLine}")
                End If
            End If
            ProgressBar1.Value += 1
        Next

        ' Extract and process potential multi-word names
        Dim multiWordMatches As MatchCollection = potentialMultiWordNameRegex.Matches(textWithoutDialogue)
        ProgressBar1.Maximum += multiWordMatches.Count
        For Each match As Match In multiWordMatches
            Dim potentialMultiWordName As String = match.Value
            Dim wordsInName As String() = potentialMultiWordName.Split(New Char() {" "}, StringSplitOptions.RemoveEmptyEntries)
            Dim allNonTitleWordsInDict As Boolean = False ' Changed logic: at least one non-title word must be in the dict
            Dim hasNonTitleWord As Boolean = False

            If wordsInName.Length > 1 Then ' Only consider as multi-word if it has more than one word
                For Each word As String In wordsInName
                    Dim cleanedWord As String = CleanName(word)
                    If Not Titles.Contains(cleanedWord, StringComparer.OrdinalIgnoreCase) Then
                        hasNonTitleWord = True
                        If Not nameDict.Contains(cleanedWord, StringComparer.OrdinalIgnoreCase) Then
                            allNonTitleWordsInDict = False ' If any non-title word is not in the dict, the flag is false
                            Exit For
                        Else
                            allNonTitleWordsInDict = True ' Set to true initially, and remains true only if all non-title words are in the dict
                        End If
                    End If
                Next

                ' Check if there was at least one non-title word and if all non-title words found are in the dictionary
                Dim cleanedMultiWordName As String = CleanName(potentialMultiWordName)
                If hasNonTitleWord AndAlso allNonTitleWordsInDict AndAlso Not names.Contains(cleanedMultiWordName, StringComparer.OrdinalIgnoreCase) Then
                    names.Add(cleanedMultiWordName)
                    IO.File.AppendAllText(Form1.Log, $"Added potential multi-word name: {cleanedMultiWordName}{vbNewLine}") ' Commented out as Form1.Log is not defined in this context
                    Console.WriteLine($"Added potential multi-word name: {cleanedMultiWordName}") ' For demonstration
                End If
            End If
            ProgressBar1.Value += 1

        Next
        StatusLbl.Text = ""
        ProgressBar1.Value = 0
        Return names
    End Function

    'Private Function ExtractNames(text As String, nameDict As List(Of String)) As List(Of String)
    '    'todo: implement function to connect aliases to names


    '    Dim names As New List(Of String)()
    '    ' Regex to find text within double quotes
    '    Dim dialogueRegex As New Regex("""([^""]*)""")
    '    ' Regex to find single words starting with a capital letter
    '    Dim singleWordNameRegex As New Regex("\b[A-Z][\w\-\']*\b")
    '    ' Regex to find sequences of one or more capitalized words
    '    Dim potentialMultiWordNameRegex As New Regex("\b[A-Z][\w\-\']+(?:\s+[A-Z][\w\-\']+)*\b")

    '    ' Remove text within double quotes to only process text outside dialogue
    '    Dim textWithoutDialogue As String = dialogueRegex.Replace(text, "")

    '    StatusLbl.Text = "Extracting names from story..."

    '    ' Extract and process single-word names
    '    Dim singleWordMatches As MatchCollection = singleWordNameRegex.Matches(textWithoutDialogue)
    '    ProgressBar1.Maximum = singleWordMatches.Count
    '    For Each match As Match In singleWordMatches
    '        Dim potentialName As String = match.Value
    '        Dim cleanedName As String = CleanName(potentialName)
    '        If Not Titles.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) Then
    '            If nameDict.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso Not CommonWords.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
    '               Not names.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) Then
    '                names.Add(cleanedName)
    '            End If
    '        End If
    '        ProgressBar1.Value += 1
    '    Next

    '    ' Extract and process potential multi-word names
    '    Dim multiWordMatches As MatchCollection = potentialMultiWordNameRegex.Matches(textWithoutDialogue)
    '    ProgressBar1.Maximum += multiWordMatches.Count
    '    For Each match As Match In multiWordMatches
    '        Dim potentialMultiWordName As String = match.Value
    '        Dim wordsInName As String() = potentialMultiWordName.Split(New Char() {" "}, StringSplitOptions.RemoveEmptyEntries)
    '        Dim allWordsInDict As Boolean = True
    '        If wordsInName.Length > 1 Then ' Only consider as multi-word if it has more than one word
    '            For Each word As String In wordsInName
    '                Dim cleanedWord As String = CleanName(word)
    '                If Not nameDict.Contains(cleanedWord, StringComparer.OrdinalIgnoreCase) Then
    '                    allWordsInDict = False
    '                    Exit For
    '                End If
    '            Next
    '            If allWordsInDict AndAlso Not names.Contains(potentialMultiWordName, StringComparer.OrdinalIgnoreCase) Then
    '                names.Add(potentialMultiWordName)
    '                IO.File.AppendAllText(Form1.Log, $"Added potential multi-word name: {potentialMultiWordName}{vbNewLine}")
    '            End If
    '        End If
    '        ProgressBar1.Value += 1
    '    Next

    '    StatusLbl.Text = ""
    '    ProgressBar1.Value = 0
    '    Return names
    'End Function
    'Private Function ExtractNames(text As String, nameDict As List(Of String)) As List(Of String)
    '    'todo: improve code to avoid false positives such as capitalized common words
    '    'todo: skip any quoted text
    '    'todo: work on finding multi-word names by checking each part against nameDict
    '    'todo: prevent listing non-capitalized words

    '    Dim names As New List(Of String)()
    '    ' Regex to find text within double quotes
    '    Dim dialogueRegex As New Regex("""([^""]*)""")
    '    ' Regex to find single words starting with a capital letter
    '    Dim singleWordNameRegex As New Regex("\b[A-Z][\w\-\']*\b")
    '    ' Regex to find sequences of one or more capitalized words
    '    ' Regex to find sequences of one or more capitalized words
    '    Dim potentialMultiWordNameRegex As New Regex("\b[A-Z][\w\-\']+(?:\s+[A-Z][\w\-\']+)*\b")

    '    ' Remove text within double quotes to only process text outside dialogue
    '    Dim textWithoutDialogue As String = dialogueRegex.Replace(text, "")

    '    StatusLbl.Text = "Extracting names from story..."

    '    Dim potentialNames As New List(Of String)()

    '    ' Process single-word names (as before, but add to potentialNames list)
    '    Dim singleWordMatches As MatchCollection = singleWordNameRegex.Matches(textWithoutDialogue)
    '    ProgressBar1.Maximum = singleWordMatches.Count
    '    For Each match As Match In singleWordMatches
    '        Dim potentialName As String = match.Value
    '        Dim cleanedName As String = CleanName(potentialName)
    '        If nameDict.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso Not CommonWords.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) AndAlso
    '           Not potentialNames.Contains(cleanedName, StringComparer.OrdinalIgnoreCase) Then
    '            potentialNames.Add(cleanedName)
    '        End If
    '        ProgressBar1.Value += 1
    '    Next

    '    ' Extract and process potential multi-word names
    '    Dim multiWordMatches As MatchCollection = potentialMultiWordNameRegex.Matches(textWithoutDialogue)
    '    ProgressBar1.Maximum += multiWordMatches.Count
    '    For Each match As Match In multiWordMatches
    '        Dim potentialMultiWordName As String = match.Value
    '        If Not potentialNames.Contains(potentialMultiWordName, StringComparer.OrdinalIgnoreCase) Then ' Avoid adding duplicates
    '            Dim wordsInName As String() = potentialMultiWordName.Split(New Char() {" "}, StringSplitOptions.RemoveEmptyEntries)
    '            If wordsInName.Length > 1 Then
    '                ' Check if at least one word is in the name dictionary and it's not entirely common words
    '                Dim foundInNameDict As Boolean = False
    '                Dim allWordsCommon As Boolean = True
    '                For Each word As String In wordsInName
    '                    Dim cleanedWord As String = CleanName(word)
    '                    If nameDict.Contains(cleanedWord, StringComparer.OrdinalIgnoreCase) Then
    '                        foundInNameDict = True
    '                        allWordsCommon = False ' If at least one is in name dict, not all are common
    '                    ElseIf Not CommonWords.Contains(cleanedWord, StringComparer.OrdinalIgnoreCase) Then
    '                        allWordsCommon = False ' If a word isn't in name dict and not common, it contributes
    '                    End If
    '                Next
    '                If foundInNameDict AndAlso Not allWordsCommon Then
    '                    potentialNames.Add(potentialMultiWordName)
    '                    IO.File.AppendAllText(Form1.Log, $"Added potential multi-word name (at least one in dict): {potentialMultiWordName}{vbNewLine}")
    '                Else ' Consider cases where all words might not be in the dict but it's still a likely name
    '                    ' Add more sophisticated checks here, like looking for common name prefixes/suffixes
    '                    ' For now, let's just add it if it's a sequence of capitalized words and not all are common
    '                    Dim allAreCommon = True
    '                    For Each word As String In wordsInName
    '                        If Not CommonWords.Contains(CleanName(word), StringComparer.OrdinalIgnoreCase) Then
    '                            allAreCommon = False
    '                            Exit For
    '                        End If
    '                    Next
    '                    If Not allAreCommon Then
    '                        ' You might want a different threshold or more specific checks here
    '                        ' For example, require at least two capitalized words if none are in the dict
    '                        If wordsInName.Length > 1 Then
    '                            ' Add this as a potential name, but you might want to flag it for lower confidence
    '                            potentialNames.Add(potentialMultiWordName)
    '                            IO.File.AppendAllText(Form1.Log, $"Added potential multi-word name (not all common): {potentialMultiWordName}{vbNewLine}")
    '                        End If
    '                    End If
    '                End If
    '            End If
    '        End If
    '        ProgressBar1.Value += 1
    '    Next

    '    StatusLbl.Text = ""
    '    ProgressBar1.Value = 0
    '    Return potentialNames.Distinct(StringComparer.OrdinalIgnoreCase).ToList() ' Ensure no duplicates
    'End Function

    Public Function GetSelectedNames() As List(Of String)
        Dim selectedNames As New List(Of String)()
        For Each item As Object In PotentialNameListBox.CheckedItems
            selectedNames.Add(item.ToString())
        Next
        Return selectedNames
    End Function

    Private Sub PotentialNameListBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles PotentialNameListBox.SelectedIndexChanged
        If PotentialNameListBox.CheckedItems.Count > 0 Then
            Form1.NextButton.Enabled = True
        Else
            Form1.NextButton.Enabled = False
        End If
    End Sub

End Class

