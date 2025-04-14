Option Explicit On
Option Strict On

'Imports System.Collections.Generic
'Imports System.IO
'Imports System.Text.RegularExpressions
'Imports FuzzySharp
'Imports java.io
'Imports opennlp.tools.chunker
'Imports opennlp.tools.langdetect
'Imports opennlp.tools.namefind
'Imports opennlp.tools.postag
'Imports opennlp.tools.sentdetect

'Public Class DialogueAttribution

'    ' --- Fields and Initialization ---
'    Private ReadOnly nerModel As NameFinderME ' OpenNLP NER model
'    Private ReadOnly langModel As LanguageDetector ' OpenNLP Language Detector
'    Private ReadOnly sentModel As SentenceDetector ' OpenNLP Sentence Detector
'    Private ReadOnly posModel As POSTaggerME ' OpenNLP POS Tagger
'    Private ReadOnly chunkModel As ChunkerME ' OpenNLP Chunker


'    ' Add other OpenNLP models as needed (POS Tagger, Parser, etc.)
'    Private ReadOnly nameDict As List(Of String) ' Your NameDict
'    Private ReadOnly AppPath As String = Application.StartupPath & "\NLP Models"
'    ' Constructor (Initialize models and NameDict)
'    Public Sub New(nameDict As List(Of String))
'        ' Load OpenNLP models (replace with your actual model loading)
'        ' Example:
'        Dim PersonFile As String = Path.Combine(AppPath, "en-ner-person.bin") ' Path to your NER model
'        If IO.File.Exists(PersonFile) Then
'            Dim nameModel As New TokenNameFinderModel(New FileInputStream(PersonFile)) ' Load into TokenNameFinderModel
'            nerModel = New NameFinderME(nameModel) ' Create NameFinderME with the model
'        Else
'            MessageBox.Show("NER model not found at: " & PersonFile, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If
'        Dim LangModelFile As String = Path.Combine(AppPath, "langdetect-183.bin") ' Path to your language model
'        If IO.File.Exists(LangModelFile) Then
'            Dim languageModel As New LanguageDetectorModel(New FileInputStream(LangModelFile)) ' Load into LanguageDetectorModel
'            langModel = New LanguageDetectorME(languageModel) ' Create LanguageDetectorME with the model
'        Else
'            MessageBox.Show("Language model not found at: " & LangModelFile, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If
'        Dim sentenceModelFile As String = Path.Combine(AppPath, "opennlp-en-ud-ewt-sentence-1.2-2.5.0.bin") ' Path to your sentence model
'        If IO.File.Exists(sentenceModelFile) Then
'            Dim sentenceModel As New SentenceModel(New FileInputStream(sentenceModelFile)) ' Load into SentenceModel
'            sentModel = New SentenceDetectorME(sentenceModel) ' Create SentenceDetectorME with the model
'        Else
'            MessageBox.Show("Sentence model not found at: " & sentenceModelFile, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If
'        'Dim posModelFile As String = Path.Combine(AppPath, "opennlp-en-ud-ewt-pos-1.2-2.5.0.bin") ' Path to your POS model
'        'If IO.File.Exists(posModelFile) Then
'        '    Dim posModel As New POSTaggerModel(New FileInputStream(posModelFile)) ' Load into POSTaggerModel
'        '    posModel = New POSTaggerME(posModel) ' Create POSTaggerME with the model
'        'Else
'        '    MessageBox.Show("POS model not found at: " & posModelFile, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        'End If
'        Dim chunkModelFile As String = Path.Combine(AppPath, "en-chunker.bin") ' Path to your chunker model
'        If IO.File.Exists(chunkModelFile) Then
'            Dim chunkModel As New ChunkerModel(New FileInputStream(chunkModelFile)) ' Load into ChunkerModel
'            chunkModel = New ChunkerME(chunkModel) ' Create ChunkerME with the model
'        Else
'            MessageBox.Show("Chunker model not found at: " & chunkModelFile, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'        End If
'        Me.nameDict = nameDict
'    End Sub

'    ' --- Main Dialogue Attribution Function ---
'    Public Function AttributeDialogue(text As String) As Dictionary(Of String, List(Of String))
'        Dim attributedDialogue As New Dictionary(Of String, List(Of String))()

'        ' 1. Extract dialogue lines (basic regex - refine with NLP later)
'        Dim dialogueLines As Dictionary(Of String, String) = ExtractDialogue(text)

'        ' 2. Process text with OpenNLP.NET (NER for character identification)
'        Dim identifiedNames As New HashSet(Of String)(StringComparer.OrdinalIgnoreCase)
'        Dim sentenceModelFile As String = "en-sent.bin" ' Replace with your sentence model file path
'        Dim sentenceModel As New SentenceModel(New FileInputStream(sentenceModelFile))
'        Dim sentenceDetector As New SentenceDetectorME(sentenceModel)
'        'Dim sentences() As String = sentenceDetector.SentenceDetect(text)

'        'For Each sentence As String In sentences
'        '    Dim sentenceNames = PerformNER(sentence) ' Implement PerformNER
'        '    For Each name As String In sentenceNames
'        '        If nameDict.Contains(name, StringComparer.OrdinalIgnoreCase) Then
'        '            identifiedNames.Add(name)
'        '        End If
'        '    Next
'        'Next

'        ' 3. Attribute dialogue lines
'        For Each key As String In dialogueLines.Keys
'            Dim dialogue As String = dialogueLines(key)
'            Dim speaker As String = key

'            If speaker.StartsWith("Unknown_") Then
'                speaker = DetermineSpeaker(text, dialogue, identifiedNames) ' Implement DetermineSpeaker
'            End If

'            ' Add dialogue to attributed dictionary
'            If Not attributedDialogue.ContainsKey(speaker) Then
'                attributedDialogue(speaker) = New List(Of String)
'            End If

'            attributedDialogue(speaker).Add(dialogue)
'        Next

'        Return attributedDialogue
'    End Function

'    ' --- Helper Functions ---
'    ' 1. Extract Dialogue (Basic - Refine with NLP)
'    Private Function ExtractDialogue(text As String) As Dictionary(Of String, String)
'        Dim dialogueLines As New Dictionary(Of String, String)()
'        Dim dialogueRegex As New Regex("(?<speaker>\b[A-Z][a-z]+(?: [A-Z][a-z]+)*\b)?(?:(?::\s*""(?<dialogue>[^""]*)"")|(?:""(?<dialogue2>[^""]*)""))", RegexOptions.IgnoreCase)

'        For Each match As Match In dialogueRegex.Matches(text)
'            Dim speaker As String = match.Groups("speaker").Value.Trim()
'            Dim dialogue As String = ""

'            If match.Groups("dialogue").Success Then
'                dialogue = match.Groups("dialogue").Value.Trim()
'            ElseIf match.Groups("dialogue2").Success Then
'                dialogue = match.Groups("dialogue2").Value.Trim()
'            End If

'            If String.IsNullOrEmpty(speaker) Then
'                ' If no speaker is directly identified, we'll assign it later
'                dialogueLines.Add("Unknown_" & dialogueLines.Count.ToString(), dialogue)
'            Else
'                dialogueLines.Add(speaker, dialogue)
'            End If
'        Next

'        Return dialogueLines
'    End Function

'    ' 2. Perform Named Entity Recognition (OpenNLP.NET)
'    Private Function PerformNER(sentence As String) As List(Of String)
'        Dim names As New List(Of String)()
'        ' Implement OpenNLP.NET NER processing here
'        ' Example (Conceptual):
'        ' Dim tokens() As String = OpenNLP.Tokenize(sentence) ' Use OpenNLP to tokenize
'        ' Dim nameSpans() As Span = nerModel.Find(tokens) ' Use NER model
'        ' For Each span As Span In nameSpans
'        '     Dim name As String = String.Join(" ", tokens.Skip(span.Start).Take(span.Length))
'        '     names.Add(name)
'        ' Next
'        Return names
'    End Function

'    ' 3. Determine Speaker (Contextual Analysis with FuzzySharp)
'    Private Function DetermineSpeaker(text As String, dialogue As String, identifiedNames As HashSet(Of String)) As String
'        Dim bestSpeaker As String = "Unknown"
'        Dim bestScore As Integer = 0

'        ' Implement contextual analysis and FuzzySharp matching here
'        ' Example (Conceptual):
'        ' Find potential speakers near the dialogue
'        ' Use FuzzySharp to score the similarity between dialogue context and speaker names
'        ' Select the speaker with the highest score

'        Return bestSpeaker
'    End Function

'    ' --- Helper Functions ---
'    ' Fuzzy Match
'    Private Function FuzzyMatch(name As String, context As String, nameDict As List(Of String)) As String
'        ' Implement FuzzySharp fuzzy matching here
'        ' Example:
'        ' Dim bestMatch As String = FuzzySharp.FindBestMatch(context, nameDict)
'        ' Return bestMatch
'        Return ""
'    End Function

'    ' Clean Name
'    Public Shared Function CleanName(name As String) As String
'        ' Remove leading/trailing whitespace and any non-alphanumeric characters
'        ' that might be common in names (e.g., hyphens, apostrophes)
'        Return Regex.Replace(name.Trim(), "[^\p{L}\p{N}\-']", "")
'    End Function
'End Class

'Imports opennlp.tools.namefind
'Imports opennlp.tools.langdetect
'Imports opennlp.tools.SentenceDetect
'Imports opennlp.tools.POS
'Imports opennlp.tools.chunker
'Imports opennlp.tools.util
'Imports System.IO
'Imports System.Windows.Forms
'Imports System.Text.RegularExpressions
'Imports FuzzySharp
'Imports opennlp.tools.tokenize
'Imports opennlp.tools.parser
'Imports java.io
'Imports opennlp.tools.postag
'Imports opennlp.tools.sentdetect

'Public Class DialogueAttribution

'    ' --- Fields and Initialization ---
'    Private ReadOnly nerModel As NameFinderME ' OpenNLP NER model
'    Private ReadOnly langModel As LanguageDetectorME ' OpenNLP Language DetectorME
'    Private ReadOnly sentModel As SentenceDetectorME ' OpenNLP Sentence DetectorME
'    Private ReadOnly posModel As POSTaggerME ' OpenNLP POS TaggerME
'    Private ReadOnly chunkModel As ChunkerME ' OpenNLP ChunkerME
'    Private ReadOnly parserModel As Parser 'OpenNLP ParserModel

'    Private ReadOnly nameDict As List(Of String) ' Your NameDict
'    Private ReadOnly AppPath As String = Application.StartupPath & "\NLP Models"

'    ' Constructor (Initialize models and NameDict)
'    Public Sub New(nameDict As List(Of String))
'        ' Load OpenNLP models
'        Try
'            Dim PersonFile As String = Path.Combine(AppPath, "en-ner-person.bin")
'            If IO.File.Exists(PersonFile) Then
'                Dim nameModel As New TokenNameFinderModel(New FileInputStream(PersonFile))
'                nerModel = New NameFinderME(nameModel)
'            Else
'                Throw New IO.FileNotFoundException("NER model not found at: " & PersonFile)
'            End If

'            Dim LangModelFile As String = Path.Combine(AppPath, "langdetect-183.bin")
'            If IO.File.Exists(LangModelFile) Then
'                Dim languageModel As New LanguageDetectorModel(New FileInputStream(LangModelFile))
'                langModel = New LanguageDetectorME(languageModel)
'            Else
'                Throw New IO.FileNotFoundException("Language model not found at: " & LangModelFile)
'            End If

'            Dim sentenceModelFile As String = Path.Combine(AppPath, "opennlp-en-ud-ewt-sentence-1.2-2.5.0.bin")
'            If IO.File.Exists(sentenceModelFile) Then
'                Dim sentenceModel As New SentenceModel(New FileInputStream(sentenceModelFile))
'                sentModel = New SentenceDetectorME(sentenceModel)
'            Else
'                Throw New IO.FileNotFoundException("Sentence model not found at: " & sentenceModelFile)
'            End If

'            Dim posModelFile As String = Path.Combine(AppPath, "opennlp-en-ud-ewt-pos-1.2-2.5.0.bin")
'            If IO.File.Exists(posModelFile) Then
'                Dim posModelData As New POSModel(New FileInputStream(posModelFile))
'                posModel = New POSTaggerME(posModelData)
'            Else
'                Throw New IO.FileNotFoundException("POS model not found at: " & posModelFile)
'            End If

'            Dim chunkModelFile As String = Path.Combine(AppPath, "en-chunker.bin")
'            If IO.File.Exists(chunkModelFile) Then
'                Dim chunkModelData As New ChunkerModel(New FileInputStream(chunkModelFile))
'                chunkModel = New ChunkerME(chunkModelData)
'            Else
'                Throw New IO.FileNotFoundException("Chunker model not found at: " & chunkModelFile)
'            End If

'            Dim parserModelFile As String = Path.Combine(AppPath, "en-parser-chunking.bin")
'            If IO.File.Exists(parserModelFile) Then
'                Dim parserModelData As New ParserModel(New FileInputStream(parserModelFile))
'                ParserModel = New Parser(parserModelData)
'            Else
'                Throw New IO.FileNotFoundException("Parser model not found at: " & parserModelFile)
'            End If

'            Me.nameDict = nameDict

'        Catch ex As io.FileNotFoundException
'            MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'            ' Consider logging the exception or taking other appropriate action
'        Catch ex As Exception
'            MessageBox.Show("Error loading models: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
'            ' Consider logging the exception or taking other appropriate action
'        End Try
'    End Sub

'    ' --- Main Dialogue Attribution Function ---
'    Public Function AttributeDialogue(text As String) As Dictionary(Of String, List(Of String))
'        Dim attributedDialogue As New Dictionary(Of String, List(Of String))()

'        ' 1. Detect sentences
'        Dim sentences() As String = DetectSentences(text)

'        ' 2. Process each sentence
'        For Each sentence As String In sentences
'            ' 3. Tokenize the sentence
'            Dim tokens() As String = TokenizeSentence(sentence)

'            ' 4. Perform NER
'            Dim names() As String = PerformNER(tokens)

'            ' 5. Perform POS tagging
'            Dim tags() As String = PerformPOSTagging(tokens)

'            ' 6. Chunk the sentence
'            'Dim chunks() As String = PerformChunking(tokens, tags)

'            ' 7. Extract dialogue and attribute it
'            ExtractAndAttributeDialogue(attributedDialogue, sentence, tokens, names)
'        Next

'        Return attributedDialogue
'    End Function

'    ' --- NLP Processing Functions ---
'    ' 1. Sentence Detection
'    Private Function DetectSentences(text As String) As String()
'        Return sentModel.SentenceDetect(text)
'    End Function

'    ' 2. Tokenization
'    Private Function TokenizeSentence(sentence As String) As String()
'        Dim tokenizer As New WhitespaceTokenizer() ' Or other tokenizer
'        Return tokenizer.tokenize(sentence)
'    End Function

'    ' 3. Named Entity Recognition (OpenNLP.NET)
'    Private Function PerformNER(tokens As String()) As String()
'        Dim names As New List(Of String)()
'        Dim nameSpans() As Span = nerModel.find(tokens)
'        For Each span As Span In nameSpans
'            Dim name As String = String.Join(" ", tokens.Skip(span.Start).Take(span.length))
'            names.Add(name)
'        Next
'        Return names.ToArray()
'    End Function

'    ' 4. Part-of-Speech Tagging
'    Private Function PerformPOSTagging(tokens As String()) As String()
'        Return posModel.Tag(tokens)
'    End Function

'    ' 5. Chunking
'    Private Function PerformChunking(tokens As String(), tags As String()) As String()
'        Return chunkModel.chunk(tokens, tags)
'    End Function

'    ' --- Dialogue Extraction and Attribution ---
'    Private Sub ExtractAndAttributeDialogue(attributedDialogue As Dictionary(Of String, List(Of String)), sentence As String, tokens As String(), names As String())
'        ' This is where the core logic for dialogue extraction and attribution goes
'        ' You'll need to:
'        ' 1. Identify dialogue within the sentence (e.g., using regex or more advanced parsing techniques)
'        ' 2. Determine the context around the dialogue
'        ' 3. Use FuzzySharp to match potential speakers (from names) to the context
'        ' 4. Add the dialogue to the attributedDialogue dictionary with the correct speaker

'        ' --- Example (Basic - Replace with your logic) ---
'        Dim dialogueRegex As New Regex("""([^""]+)""") ' Basic regex to find text within quotes
'        Dim matches As MatchCollection = dialogueRegex.Matches(sentence)

'        For Each match As Match In matches
'            Dim dialogue As String = match.Value.Trim("""")
'            Dim speaker As String = DetermineSpeaker(sentence, dialogue, names) ' Determine Speaker

'            If Not attributedDialogue.ContainsKey(speaker) Then
'                attributedDialogue(speaker) = New List(Of String)()
'            End If
'            attributedDialogue(speaker).Add(dialogue)
'        Next
'    End Sub

'    ' --- Speaker Determination ---
'    Private Function DetermineSpeaker(sentence As String, dialogue As String, names As String()) As String
'        ' This is where you implement the logic to determine who said the dialogue
'        ' You'll need to use FuzzySharp to compare the dialogue and its context to the names

'        ' --- Example (Basic - Replace with your logic) ---
'        If names.Length > 0 Then
'            ' Use FuzzySharp to find the best match from names based on the sentence
'            Dim bestMatch = FuzzySharp.Fuzz.Choices(dialogue, names)
'            Return bestMatch.First().Value
'        Else
'            Return "Unknown"
'        End If
'    End Function

'    ' --- Helper Functions ---
'    ' Fuzzy Match
'    Private Function FuzzyMatch(name As String, context As String) As Integer
'        ' Implement FuzzySharp fuzzy matching here
'        Return FuzzySharp.Fuzz.Ratio(name, context)
'    End Function

'    ' Clean Name
'    Public Shared Function CleanName(name As String) As String
'        ' Remove leading/trailing whitespace and any non-alphanumeric characters
'        ' that might be common in names (e.g., hyphens, apostrophes)
'        Return Regex.Replace(name.Trim(), "[^\p{L}\p{N}\-']", "")
'    End Function
'End Class