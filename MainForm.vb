' Biniam Asnake's Indexing System
' Programmed by: Biniam Asnake
' ID No: GSR/0809/03
' 
' Submitted to: Dr. Dereje Teferi (PhD)
' Date: May, 2011

Imports System.IO

Public Class frmIRS

#Region "Declaration of Global Variables"

    'Declare three string variables 
    ' 1. filename1 -> to store the name of the file
    ' 2. content1 -> to store the content of the file
    ' 3. newContent1 -> to store the new content1 where the punctuations are removed
    ' Dim filename1, content1, newContent1 As String

    'To store the folder name where all the files reside (the document corpus)
    Dim folderName As String

    Dim allFiles As String()

    'Dim filename1 As String = "G:\2nd Semester\2. Information Storage & Retrieval\Assignments\B IRS\B IRS\sample.txt"
    Dim filename1 As String = String.Empty
    Dim filename2 As String = String.Empty
    Dim filename3 As String = String.Empty
    Dim filename4 As String = String.Empty
    Dim filename5 As String = String.Empty

    Dim content1 As String = String.Empty
    Dim content2 As String = String.Empty
    Dim content3 As String = String.Empty
    Dim content4 As String = String.Empty
    Dim content5 As String = String.Empty

    Dim newContent As String

    'Declare a String array to that stores ... 
    Dim stringArray As String()               'each of the tokenized words

    ' Dim myNewStringArray As String()            'the final list of words that are free from stopwords
    Dim myNewStringArray1 As String()
    Dim myNewStringArray2 As String()
    Dim myNewStringArray3 As String()
    Dim myNewStringArray4 As String()
    Dim myNewStringArray5 As String()

    'Decalre an Integer to store m -> the number of times VC repeats
    Dim m As Integer

    'Declare an Integer to store Term Frequency
    Dim termFreq As Double

    'Declare an Integer to store Document Frequency
    Dim docFreq As Integer

    'Declare an Integer to store Inverse Document Frequency
    Dim invDocFreq As Double

    'Declare an Integer to store Term Weighting (tf*idf)
    Dim tfidf As Double

#End Region

    Private Sub frmIRS_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Call stopListReader("C:\Users\Biniam\Desktop\B IRS\B IRS\stoplist.txt")

    End Sub

    Private Sub btnBrowse1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse1.Click

        ''Show the open dialog and if the user clicks the open button load the file
        'If OpenFileDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
        '    Try
        '        'Save the file name in the variable called "filename1"
        '        filename1 = OpenFileDialog1.FileName
        '        txtFile1.Text = filename1
        '    Catch ex As Exception               'FileNotFoundException
        '        'If there is any exception related to opening file, display an exception (error) message to the user
        '        MessageBox.Show("Please select a text file", "B IRS", MessageBoxButtons.OK, MessageBoxIcon.Error)
        '    End Try
        'End If

        'Show the open dialog and if the user clicks the open button load the file
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then
            Try
                'Save the file name in the variable called "filename1"
                folderName = FolderBrowserDialog1.SelectedPath
                txtFile1.Text = folderName
            Catch ex As Exception               'FileNotFoundException
                'If there is any exception related to opening file, display an exception (error) message to the user
                MessageBox.Show("Please select a folder where the documents reside.", "B IRS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End If

        allFiles = IO.Directory.GetFiles(folderName, "*.txt")

        btnStem.Enabled = True

        'OUTPUT OUTPUT OUTPUT OUTPUT OUTPUT OUTPUT
        'For k = 0 To allFiles.Length - 1              'allFiles.Length - 1

        'filename1 = allFiles(k)

        'If k = 0 Then
        'filename1 = allFiles(k)
        'End If

        'If k = 1 Then
        '    filename2 = allFiles(k)
        'End If

        'If k = 2 Then
        '    filename3 = allFiles(k)
        'End If

        'If k = 3 Then
        '    filename4 = allFiles(k)
        'End If

        'If k = 4 Then
        '    filename5 = allFiles(k)
        'End If
        ' MsgBox("Output:  " & allFiles(k))
        'Next

    End Sub

#Region "Declaration of Global Variables #2"

    'Declare a StreamReader object
    Dim swrStreamReader As StreamReader
    Dim stopListItem As Object
    Dim stopword As String
    Dim stopWordArray As String()

#End Region
    
    Private Sub stopListReader(ByVal filename As String)

        'Read from the file
        Try
            swrStreamReader = IO.File.OpenText(filename)
            'Save the whole content of the file in the variable called "content1"
            stopListItem = swrStreamReader.ReadLine

            While Not stopListItem = Nothing

                stopword = stopword & stopListItem.ToString.Trim & " "
                stopListItem = swrStreamReader.ReadLine

            End While

            'When finished reading from the file, close the StreamReader
            swrStreamReader.Close()

            Dim separators As Char()
            separators = {" ", vbNewLine}

            'Split the document into single words
            stopword = stopword.Trim.ToLower

            stopWordArray = stopword.Split(separators)

        Catch ex As Exception
            MessageBox.Show("The stop list file is not loaded! Please correct this problem.", "B IRS", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End
        End Try

    End Sub

    Private Function stopWordRemover(ByVal content As String()) As String()

        Dim i As Integer
        Dim con As String = String.Empty
        Dim myNewStringArray As String() = Nothing

        For i = 0 To stringArray.Length - 1

            If stopWordArray.Contains(stringArray(i)) = False And stringArray(i).Equals("") = False And IsNumeric(stringArray(i)) = False Then
                con = con & stringArray(i) & " "
            End If

        Next

        If con.Equals("") = False Then
            myNewStringArray = con.Trim.Split(" ")
        End If

        Return myNewStringArray

    End Function

    Private Sub btnStem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStem.Click
        If allFiles.Length <> 0 Then

            For k = 0 To allFiles.Length - 1

                filename1 = allFiles(k)

                If filename1.Equals(String.Empty) = False Then

                    content1 = ReadFromFile(filename1)

                    Call Normalizer(content1)

                    stringArray = tokenizer(content1)

                    'Function call and Return
                    myNewStringArray1 = stopWordRemover(stringArray)

                    Call Stemmer(myNewStringArray1)

                    Call WriteToFile(filename1, myNewStringArray1)


                End If
            Next

            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim myStreamReader As StreamReader

            Dim old As String = String.Empty

            If IO.File.Exists(tempFileName) Then
                myStreamReader = IO.File.OpenText(tempFileName)

                'Declare a string variable to store the content that is read from the file
                Dim content As String = String.Empty

                'Save the whole content of the file in the variable called "content"
                Dim item As Object
                item = myStreamReader.ReadLine

                While Not item = Nothing

                    content = item.ToString.Trim

                    myNewStringArray5 = content.Split(" ")

                    Dim k As Integer = 0

                    While k < myNewStringArray5.Length

                        If old.Contains(myNewStringArray5(k)) Then
                            k += 1
                            Continue While

                        Else
                            termFreq = TermFrequency(myNewStringArray5(k), myNewStringArray5)
                            docFreq = DocumentFrequency(myNewStringArray5(k))
                            invDocFreq = InverseDocumentFrequency(docFreq)
                            tfidf = TermWeighting(termFreq, invDocFreq)

                            Call WriteToFile(myNewStringArray5(k), termFreq, docFreq, invDocFreq, tfidf)
                            old &= myNewStringArray5(k).Trim + " "
                        End If
                        k += 1
                    End While

                    item = myStreamReader.ReadLine

                End While

                'When finished reading from the file, close the StreamReader
                myStreamReader.Close()

            End If

            'Finally
            If Not filename1.Equals(String.Empty) Or filename2.Equals(String.Empty) = False Or filename3.Equals(String.Empty) = False Or filename4.Equals(String.Empty) = False Or filename5.Equals(String.Empty) = False Then

                MessageBox.Show("The vocabulary is saved to the file Successfully.", "B IRS", MessageBoxButtons.OK, MessageBoxIcon.Information)
                btnOpenOutputFile.Enabled = True

            End If

        Else

            MessageBox.Show("There is no text file in the folder you specified.", "B IRS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

        End If

        btnStem.Enabled = False


    End Sub

    Private Function ReadFromFile(ByVal filename As String) As String

        'Declare a StreamReader object and read from the file
        Dim myStreamReader As StreamReader = Nothing

        If IO.File.Exists(filename) Then
            myStreamReader = IO.File.OpenText(filename)
        
            'Declare a string variable to store the content that is read from the file
        Dim content As String = String.Empty

        'Save the whole content of the file in the variable called "content"
        Dim item As Object
        item = myStreamReader.ReadLine

        While Not item = Nothing

                content = content & item.ToString.Trim & vbNewLine
                item = myStreamReader.ReadLine

        End While

        'When finished reading from the file, close the StreamReader
        myStreamReader.Close()

            Return content

        Else
            MessageBox.Show("The file doesnot exist." & vbCrLf & "Please make sure the file is avialable in the location you specified.", "B IRS", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Function

    Private Sub Stemmer(ByVal myNewStringArray As String())

        If IsNothing(myNewStringArray) = False Then
            For k = 0 To myNewStringArray.Length - 1

                stepOne(myNewStringArray(k))
                stepTwo(myNewStringArray(k))
                stepThree(myNewStringArray(k))
                stepFour(myNewStringArray(k))
                stepFive(myNewStringArray(k))

            Next
        End If

    End Sub

    Dim max As Integer = 0      'The maximum number is all terms that are mentioned in the text of the document i

    Private Function TermFrequency(ByVal word As String, ByVal myNewStringArray2 As String()) As Double

        Dim i As Integer
        Dim freq As Integer = 0
        Dim tf As Double = 0

        'Declare a string variable to store the content that is read from the file
        Dim content As String = String.Empty

        Dim myStreamReader As StreamReader

        If IO.File.Exists(tempFileName) Then
            myStreamReader = IO.File.OpenText(tempFileName)

            'Save the whole content of the file in the variable called "content"
            Dim item As Object
            item = myStreamReader.ReadLine

            While Not item = Nothing

                content = item.ToString.Trim & vbNewLine

                'myNewStringArray1 = content.ToString.Trim.Split(" ")
                myNewStringArray1 = content.ToString.Trim.Split(vbNewLine)

                max = myNewStringArray1.Length()

                For i = 0 To myNewStringArray1.Length - 1

                    If word.Trim.Equals(myNewStringArray1(i).Trim) Then
                        freq += 1
                    End If

                Next

                item = myStreamReader.ReadLine

            End While

            'When finished reading from the file, close the StreamReader
            myStreamReader.Close()

            tf = Format(freq / max, "0.0000")
            Return tf

        Else
            MessageBox.Show("The Document doesnot contain any term or all the terms are stopwords", "B IRS", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If

    End Function

    Private Function DocumentFrequency(ByVal word As String) As Integer

        Dim df As Integer = 0

        'Declare a StreamReader object and read from the file
        Dim myStreamReader As StreamReader = Nothing

        If txtFile1.Text.Equals(String.Empty) = False Then

            If IO.File.Exists(tempFileName) Then
                myStreamReader = IO.File.OpenText(tempFileName)

                'Declare a string variable to store the content that is read from the file
                Dim content As String = String.Empty

                'Save the whole content of the file in the variable called "content"
                Dim item As Object
                item = myStreamReader.ReadLine

                'Dim mySeparators As Char() = {" "}

                While Not item = Nothing

                    content = item.ToString.Trim & vbNewLine
                    myNewStringArray1 = content.ToString.Trim.Split(" ")

                    If myNewStringArray1.Contains(word.Trim) Then
                        df += 1
                    End If

                    item = myStreamReader.ReadLine

                End While

                'When finished reading from the file, close the StreamReader
                myStreamReader.Close()

            End If

        End If

        Return df

    End Function

    Private Function InverseDocumentFrequency(ByVal df As Integer) As Double

        Dim freq As Integer = 0
        Dim idf As Double = 0

        'max = allFiles.Length
        'MsgBox(max)
        Try
            freq = max / df
            idf = Format(Math.Log(freq, 2), "0.0000")
        Catch ex As Exception

        End Try

        Return idf

    End Function

    Private Function TermWeighting(ByVal tf As Double, ByVal idf As Double) As Double

        Dim termWeight As Double = 0

        termWeight = Format(tf * idf, "0.0000")

        Return termWeight

    End Function

    Dim tempFileName As String

    Private Sub WriteToFile(ByVal filename As String, ByVal myNewStringArray As String())

        'Finally save the stopword removed, tokenized, normalized and stemmed vocabulary to a temporary file

        If IsNothing(myNewStringArray) = False Then

            Try
                tempFileName = folderName & "\Temp.txt"

            Catch ex As Exception
                tempFileName = "C:\Temp.txt"
            End Try

            'Declare a StreamReader object and read from the file
            Dim myStreamWriter As StreamWriter

            If IO.File.Exists(tempFileName) Then
                myStreamWriter = IO.File.AppendText(tempFileName)
            Else
                myStreamWriter = IO.File.CreateText(tempFileName)
            End If

            'Save the whole content of the file in the variable called "content1"
            Dim outputToFile As String = String.Empty

            For k = 0 To myNewStringArray.Length - 1

                outputToFile = outputToFile.Trim & " " & myNewStringArray(k).Trim

            Next

            'Put differentiator to the different docs
            outputToFile = outputToFile.Trim & vbNewLine

            'Finally, write the output to the new file
            myStreamWriter.WriteLine(outputToFile.Trim)
            myStreamWriter.Close()

        End If

    End Sub

    Dim newFileName As String = String.Empty

    Private Sub WriteToFile(ByVal word As String, ByVal tf As Double, ByVal df As Integer, ByVal idf As Double, ByVal tfidf As Double)

        'Finally save the stopword removed, tokenized, normalized and stemmed vocabulary to a file
        'by the name 'Index File.txt'

        If word.Equals(String.Empty) = False Then

            Try
                newFileName = folderName & "\Index File.txt"
                'MsgBox(newFileName)
            Catch ex As Exception
                newFileName = "C:\Index File.txt"
            End Try

            'Save the whole content of the file in the variable called "content1"
            Dim outputToFile As String = String.Empty

            'newFileName = filename.Substring(0, filename.Length - 4) & "_new" & filename.Substring(filename.Length - 4)
            '        MsgBox(newFileName)

            'Declare a StreamReader object and read from the file
            Dim myStreamWriter As StreamWriter

            If IO.File.Exists(newFileName) Then
                myStreamWriter = IO.File.AppendText(newFileName)
            Else
                myStreamWriter = IO.File.CreateText(newFileName)
                outputToFile = "KEYWORDS" & vbTab & vbTab & "|" & "  TERM FREQUENCY" & vbTab & "|" & "  DOCUMENT FREQUENCY" & vbTab & "|" & "  INV DOC FREQUENCY" & vbTab & "|" & "  TERM WEIGHTING" & vbCrLf
                outputToFile &= "==================================================================================================================="
                myStreamWriter.WriteLine(outputToFile)
            End If

            outputToFile = word.Trim & vbTab & vbTab & vbTab & "|" & vbTab & tf & vbTab & vbTab & "|" & vbTab & df & vbTab & vbTab & "|" & vbTab & idf & vbTab & vbTab & "|" & vbTab & tfidf & vbCrLf

            'Finally, write the output to the new file
            myStreamWriter.WriteLine(outputToFile.Trim)
            myStreamWriter.Close()

        End If

    End Sub

    Private Sub Normalizer(ByRef content As String)

        'NORMALIZATION - Change the case of the words
        content = content.Trim.ToLower

    End Sub

    Private Function tokenizer(ByRef content As String) As String()

        'Remove any punctuation mark
        Call punctuationRemover(content)

        'Split the document into single words
        stringArray = content.Trim.Split(" ")

        Return stringArray

    End Function

    Private Sub punctuationRemover(ByRef content As String)

        'Declare and Initialize an array that will store the list of punctuation marks in English Language
        Dim punctuations() As String = {"?", ",", ".", "<", ">", ";", ":", "{", "}", "|", "[", "]", "!", "&", "*", "`", "~", "-"}

        For i = 0 To punctuations.Length - 1

            content = content.Replace(punctuations(i), " ")

        Next

    End Sub

    Private Function isVowel(ByVal check As Char, ByVal number As Integer) As Boolean

        If check.ToString.ToLower = "a" Or
            check.ToString.ToLower = "e" Or
            check.ToString.ToLower = "i" Or
            check.ToString.ToLower = "o" Or
            check.ToString.ToLower = "u" Or
            check.ToString.ToLower = "y" And number = 1 Then

            'If the character is a Vowel, return True
            Return True

        Else
            'If the character is a consonant, return False
            Return False

        End If

    End Function

    Private Function calcM(ByVal word As String) As Integer

        'Declare a Integer variable to that stores the value to be returned
        Dim myInt As Integer
        Dim i As Integer

        For i = 0 To word.Length - 2        ' because the last letter need not to be checked
            If isVowel(word(i), word.Length - 1) = True Then

                'As long as 'i' is not the last character, check if the next character is a consonant
                If (i + 1) <= word.Length - 1 Then

                    If isVowel(word(i + 1), word.Length - 1) = False Then
                        myInt += 1
                        i += 1          ' Now, the letter @ i is known that it is constant. So, no need to check it again if it is a vowel.
                    End If

                End If

            End If
        Next

        Return myInt

    End Function

    Private Sub stepOne(ByRef word As String)

        'Step1 deals with plurals and past participles.

        'Step 1a

        If word.EndsWith("sses") Then
            'sses -> ss
            word = word.Substring(0, word.Length - 2)

        ElseIf word.EndsWith("ies") Then
            'ies -> i
            word = word.Substring(0, word.Length - 2)

        ElseIf word.EndsWith("ss") Then
            'ss -> ss
            word = word.Substring(0)

        ElseIf word.EndsWith("s") Then
            's -> 
            word = word.Substring(0, word.Length - 1)
        End If


        'Step 1b
        Dim rules1b As Boolean = False

        If word.EndsWith("eed") And word.Length > 0 AndAlso calcM(word.Substring(0, word.Length - 3)) > 0 Then
            '(m > 0) eed -> ee
            ' Remeber: calcM() function returns an integer value
            word = word.Substring(0, word.Length - 1)
            Exit Sub
        ElseIf word.EndsWith("ed") And word.Length > 0 AndAlso containsVowel(word.Substring(0, word.Length - 2)) = True Then
            '(*v*) ed -> 
            word = word.Substring(0, word.Length - 2)
            ' if the second or third of the rules in Step 1b is successful, the following is done:
            rules1b = True
        ElseIf word.EndsWith("ing") And word.Length > 0 AndAlso containsVowel(word.Substring(0, word.Length - 3)) = True Then
            '(*v*) ing -> 
            word = word.Substring(0, word.Length - 3)
            ' if the second or third of the rules in Step 1b is successful, the following is done:
            rules1b = True
        End If

        If rules1b = True Then
            ' If any of the above two statements are correct, check the following

            If word.EndsWith("at") Then
                'at -> ate
                word = word.Insert(word.Length, "e")
            ElseIf word.EndsWith("bl") Then
                'bl -> ble
                word = word.Insert(word.Length, "e")
            ElseIf word.EndsWith("iz") Then
                'iz -> ize
                word = word.Insert(word.Length, "e")

            ElseIf endsWithDoubleConsonant(word) And Not (word.EndsWith("l") Or word.EndsWith("s") Or word.EndsWith("z")) Then
                '(*d and not(*L or *S or *Z)) -> single letter
                word = word.Substring(0, word.Length - 1)

            ElseIf endsWithCVC(word) AndAlso calcM(word) = 1 Then
                '(m = 1 and *o) -> E
                word = word.Insert(word.Length, "e")

            End If

        End If

        ' Step 1c
        If word.EndsWith("y") And word.Length > 0 AndAlso containsVowel(word.Substring(0, word.Length - 1)) = True Then
            '(*v*) Y -> i
            word = word.Substring(0, word.Length - 1)
            word = word.Insert(word.Length, "i")
        End If

    End Sub

    Private Sub stepTwo(ByRef word As String)

        If word.EndsWith("ational") AndAlso calcM(word.Substring(0, word.Length - 7)) > 0 Then
            '(m > 0) ational -> ate
            word = word.Substring(0, word.Length - 7)
            word = word.Insert(word.Length, "ate")
        ElseIf word.EndsWith("tional") AndAlso calcM(word.Substring(0, word.Length - 6)) > 0 Then
            '(m > 0) tional -> tion
            word = word.Substring(0, word.Length - 6)
            word = word.Insert(word.Length, "tion")
        ElseIf word.EndsWith("enci") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) enci -> ence
            word = word.Substring(0, word.Length - 4)
            word = word.Insert(word.Length, "ence")
        ElseIf word.EndsWith("anci") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) anci -> ance
            word = word.Substring(0, word.Length - 4)
            word = word.Insert(word.Length, "ance")
        ElseIf word.EndsWith("izer") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) izer -> ize
            word = word.Substring(0, word.Length - 4)
            word = word.Insert(word.Length, "ize")
        ElseIf word.EndsWith("abli") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) abli -> able
            word = word.Substring(0, word.Length - 4)
            word = word.Insert(word.Length, "able")
        ElseIf word.EndsWith("alli") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) alli -> al
            word = word.Substring(0, word.Length - 4)
            word = word.Insert(word.Length, "al")
        ElseIf word.EndsWith("entli") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) entli -> ance
            word = word.Substring(0, word.Length - 5)
            word = word.Insert(word.Length, "ent")
        ElseIf word.EndsWith("eli") AndAlso calcM(word.Substring(0, word.Length - 3)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 3)
            word = word.Insert(word.Length, "e")
        ElseIf word.EndsWith("ousli") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) ousli -> ous
            word = word.Substring(0, word.Length - 5)
            word = word.Insert(word.Length, "ous")
        ElseIf word.EndsWith("ization") AndAlso calcM(word.Substring(0, word.Length - 7)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 7)
            word = word.Insert(word.Length, "ize")
        ElseIf word.EndsWith("ation") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 5)
            word = word.Insert(word.Length, "ate")
        ElseIf word.EndsWith("ator") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 4)
            word = word.Insert(word.Length, "ate")
        ElseIf word.EndsWith("alism") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 5)
            word = word.Insert(word.Length, "al")
        ElseIf word.EndsWith("iveness") AndAlso calcM(word.Substring(0, word.Length - 7)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 7)
            word = word.Insert(word.Length, "ive")
        ElseIf word.EndsWith("fulness") AndAlso calcM(word.Substring(0, word.Length - 7)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 7)
            word = word.Insert(word.Length, "ful")
        ElseIf word.EndsWith("ousness") AndAlso calcM(word.Substring(0, word.Length - 7)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 7)
            word = word.Insert(word.Length, "ous")
        ElseIf word.EndsWith("aliti") AndAlso calcM(word.Substring(0, word.Length - 6)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 6)
            word = word.Insert(word.Length, "al")
        ElseIf word.EndsWith("iviti") AndAlso calcM(word.Substring(0, word.Length - 6)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 6)
            word = word.Insert(word.Length, "ive")
        ElseIf word.EndsWith("biliti") AndAlso calcM(word.Substring(0, word.Length - 6)) > 0 Then
            '(m > 0) eli -> e
            word = word.Substring(0, word.Length - 6)
            word = word.Insert(word.Length, "ble")
        End If

    End Sub

    Private Sub stepThree(ByRef word As String)

        If word.EndsWith("icate") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) ational -> ate
            word = word.Substring(0, word.Length - 5)
            word = word.Insert(word.Length, "ic")
        ElseIf word.EndsWith("ative") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) tional -> tion
            word = word.Substring(0, word.Length - 5)
        ElseIf word.EndsWith("alize") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) enci -> ence
            word = word.Substring(0, word.Length - 5)
            word = word.Insert(word.Length, "al")
        ElseIf word.EndsWith("iciti") AndAlso calcM(word.Substring(0, word.Length - 5)) > 0 Then
            '(m > 0) anci -> ance
            word = word.Substring(0, word.Length - 5)
            word = word.Insert(word.Length, "ic")
        ElseIf word.EndsWith("ical") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) izer -> ize
            word = word.Substring(0, word.Length - 4)
            word = word.Insert(word.Length, "ic")
        ElseIf word.EndsWith("ful") AndAlso calcM(word.Substring(0, word.Length - 3)) > 0 Then
            '(m > 0) abli -> able
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("ness") AndAlso calcM(word.Substring(0, word.Length - 4)) > 0 Then
            '(m > 0) alli -> al
            word = word.Substring(0, word.Length - 4)
        End If

    End Sub

    Private Sub stepFour(ByRef word As String)

        If word.EndsWith("al") AndAlso calcM(word.Substring(0, word.Length - 2)) > 1 Then
            '(m > 0) al -> 
            word = word.Substring(0, word.Length - 2)
        ElseIf word.EndsWith("ance") AndAlso calcM(word.Substring(0, word.Length - 4)) > 1 Then
            '(m > 0) ance -> 
            word = word.Substring(0, word.Length - 4)
        ElseIf word.EndsWith("ence") AndAlso calcM(word.Substring(0, word.Length - 4)) > 1 Then
            '(m > 0) ence -> 
            word = word.Substring(0, word.Length - 4)
        ElseIf word.EndsWith("er") AndAlso calcM(word.Substring(0, word.Length - 2)) > 1 Then
            '(m > 0) er -> 
            word = word.Substring(0, word.Length - 2)
        ElseIf word.EndsWith("ic") AndAlso calcM(word.Substring(0, word.Length - 2)) > 1 Then
            '(m > 0) tional -> tion
            word = word.Substring(0, word.Length - 2)
        ElseIf word.EndsWith("able") AndAlso calcM(word.Substring(0, word.Length - 4)) > 1 Then
            '(m > 0) able -> 
            word = word.Substring(0, word.Length - 4)
        ElseIf word.EndsWith("ible") AndAlso calcM(word.Substring(0, word.Length - 4)) > 1 Then
            '(m > 0) ible -> 
            word = word.Substring(0, word.Length - 4)
        ElseIf word.EndsWith("ant") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) ant -> 
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("ement") AndAlso calcM(word.Substring(0, word.Length - 5)) > 1 Then
            '(m > 0) ement -> 
            word = word.Substring(0, word.Length - 5)
        ElseIf word.EndsWith("ment") AndAlso calcM(word.Substring(0, word.Length - 4)) > 1 Then
            '(m > 0) ment -> 
            word = word.Substring(0, word.Length - 4)
        ElseIf word.EndsWith("ent") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) ent ->
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("ion") And
              (word.EndsWith(word.Length - 4).Equals("s") Or word.EndsWith(word.Length - 4).Equals("t")) AndAlso calcM(word.Substring(0, word.Length - 4)) > 1 Then
            '(m > 1) and (*S or *T) ion -> 
            'The meaning of the above code is:
            ' if m is greater than 1 and if the word ends with 'ion' and if the letter before the 'ion' is 's' or 't'
            word = word.Substring(0, word.Length - 4)
        ElseIf word.EndsWith("ou") AndAlso calcM(word.Substring(0, word.Length - 2)) > 1 Then
            '(m > 0) ou -> 
            word = word.Substring(0, word.Length - 2)
        ElseIf word.EndsWith("ism") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) ism -> 
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("ate") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) ate -> 
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("iti") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) iti -> 
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("ous") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) ous -> 
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("ive") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) ive -> 
            word = word.Substring(0, word.Length - 3)
        ElseIf word.EndsWith("ize") AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 0) ize -> 
            word = word.Substring(0, word.Length - 3)
        End If

    End Sub

    Private Sub stepFive(ByRef word As String)

        'Step 5a

        If word.EndsWith("e") AndAlso calcM(word.Substring(0, word.Length - 1)) > 1 Then
            '(m > 1) E ->
            word = word.Substring(0, word.Length - 1)
        ElseIf word.EndsWith("e") And word.Length > 0 AndAlso Not endsWithCVC(word.Substring(0, word.Length - 1)) AndAlso calcM(word.Substring(0, word.Length - 1)) > 1 Then
            '(m > 1 and not *o) E -> 
            word = word.Substring(0, word.Length - 1)
        End If

        'Step 5b

        If word.EndsWith("e") And word.Length > 0 AndAlso endsWithDoubleConsonant(word.Substring(0, word.Length - 1)) AndAlso calcM(word.Substring(0, word.Length - 3)) > 1 Then
            '(m > 1 and *d and *L) -> Single Letter
            word = word.Substring(0, word.Length - 1)
        End If

    End Sub

    Private Function containsVowel(ByVal word As String) As Boolean

        Dim contains As Boolean = False
        Dim i As Integer

        For i = 0 To word.Length - 2        ' Here I made a modification to Porter's Algotithm

            'Explanation
            'PORTER     'Iterate through all the characters and check if any of them is a vowel [word.Length - 1]
            'Problem - If we enter for example "feed", the output is "fe" only -> which is an incorrect answer
            'BINIAM A.     'Iterate through all the characters except the last character and check if any of them is a vowel    [word.Length - 2]
            'Solution - If we enter for example "feed", the output is "feed" -> which is the correct answer

            If isVowel(word(i), word.Length - 1) = True Then       'And word(i) <> word(word.Length - 1) = False

                contains = True

            End If

        Next

        Return contains

    End Function

    Private Function endsWithDoubleConsonant(ByVal word As String) As Boolean

        Dim endsWithConsonants As Boolean = False
        Dim i, j, startingPoint As Integer

        'MsgBox(word)
        startingPoint = word.Length - 1

        For i = startingPoint To 0 Step -1

            j = i - 1
            ' If the last number is the same as it's predicesor
            If j > 0 AndAlso word(i) = word(j) Then
                If isVowel(word(i), word.Length - 1) = False Then
                    endsWithConsonants = True
                End If
            End If
        Next

        Return endsWithConsonants

    End Function

    Private Function endsWithCVC(ByVal word As String) As Boolean

        Dim endsCVC As Boolean = False
        Dim i As Integer

        If word.Length >= 3 Then

            For i = word.Length - 1 To 0 Step -1

                If isVowel(word(word.Length - 3), word.Length - 1) = False AndAlso
                   isVowel(word(word.Length - 2), word.Length - 1) = True AndAlso
                   (isVowel(word(word.Length - 1), word.Length - 1) = False AndAlso
                            Not (word.EndsWith("W") Or word.EndsWith("X") Or word.EndsWith("Y"))) Then
                    endsCVC = True
                End If
            Next

        End If

        Return endsCVC

    End Function

    Private Sub btnExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

        End

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        End

    End Sub

    Private Sub btnOpenOutputFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpenOutputFile.Click

        Dim indexFile As String = String.Empty
        indexFile = ReadFromFile(newFileName)

        frmOutput.RichTextBox1.Text = indexFile

        frmOutput.Show()

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click

        AboutBox1.Show()

    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewToolStripMenuItem.Click

        btnStem_Click(Nothing, Nothing)

    End Sub

    Private Sub BrowseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BrowseToolStripMenuItem.Click

        btnBrowse1_Click(Nothing, Nothing)

    End Sub

End Class