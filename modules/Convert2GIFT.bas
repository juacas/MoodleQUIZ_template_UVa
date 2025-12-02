' The MIT License
' Copyright (c) 2005 Mikko Rusama
'
' Permission is hereby granted, free of charge, to any person obtaining a copy of this
' software and associated documentation files (the "Software"), to deal in the Software
' without restriction, including without limitation the rights to use, copy, modify, merge,
' publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons
' to whom the Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all copies or
' substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
' BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
' NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
' DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
'
' GIFT Converter
' Version 1.0, Updated 24.05.2020
' Author: Juan Pablo de Castro (juan.pablo.de.castro@gmail.com)
' Version 0.8, Updated 8.10.2004
' Author: Mikko Rusama (mikko.rusama@iki.fi)
'
' A macro for converting a Word document with questions to the native GIFT questionnaire
' format supported by Moodle (www.moodle.org). Questions are defined as different
' Word styles; style definitions are below.
'
' Supported question types are:
'  1. Multiple Choice Question
'  2. Matching Question
'  3. Short Answer Question
'  4. True-False Question (statements)
'  5. Numerical Question
'  6. Missing Word Question (only 1 right answer supported)
'
' Question feedback is also supported as well as weighted answers for Multiple Choice
' Questions.
'
' Copyright 2004- SoberIT, Helsinki University of Technology
' Copytight 2020- Eduvalab, University of Valladolid
' Changes:
'  8.10.2004 Fixed decimal converter bug. Bug replaced commas with dots in the question choices.
'  24.05.2020 Single Choice Question



'********************************************************
' Style definitions. The styles defined below are used in the conversion.
'********************************************************

' General purpose styles.
Const STYLE_FEEDBACK = "Feedback"
Const STYLE_ANSWERWEIGHT = "AnswerWeight"
Const STYLE_NORMAL = "Normal"

' Styles for multiple and single choice questions
Const STYLE_SINGLECHOICEQ = "SingleChoiceQ"
Const STYLE_MULTIPLECHOICEQ = "MultipleChoiceQ"
Const STYLE_RIGHT_ANSWER = "RightAnswer"
Const STYLE_WRONG_ANSWER = "WrongAnswer"
Const STYLE_UNMARKED_ANSWER = "Unmarked"

' Styles for matching pair questions
Const STYLE_MATCHINGQ = "MatchingQ"
Const STYLE_LEFT_PAIR = "LeftPair"
Const STYLE_RIGHT_PAIR = "RightPair"

' Styles for true-false questions
Const STYLE_TRUESTATEMENT = "TrueStatement"
Const STYLE_FALSESTATEMENT = "FalseStatement"

' Style for short answer question
Const STYLE_SHORTANSWERQ = "ShortAnswerQ"

' Style for numerical question
Const STYLE_NUMERICALQ = "NumericalQ"

' Style for missing word question
Const STYLE_MISSINGWORDQ = "MissingWordQ"
Const STYLE_BLANK_WORD = "BlankWord"

'********************************************************
' General constants and variable definitions
'********************************************************
' saves the current question type
Dim questionType As String

' Prefix for the filename
Const FILE_PREFIX = "Preguntas_Moodle_"

Const COPYRIGHT = "Copyright 2004 SoberIT, Helsinki University of Technology" & vbCr _
                & "Copytight 2020 EDUVALAB, University of Valladolid"

'********************************************************
' GIFT question tags
'********************************************************
Const TAG_QUESTION_START = " {"
Const TAG_QUESTION_END = "}"
Const TAG_TRUE_CHOICE = "{T}"
Const TAG_FALSE_CHOICE = "{F}"
Const TAG_RIGHT_ANSWER = "="
Const TAG_UNMARKED_ANSWER = "~%0%"
Const TAG_RIGHT_NUMERICAL_ANSWER = "#"
Const TAG_WRONG_ANSWER = "~"
Const TAG_MATCHINGQ_ARROW = " -> "
Const TAG_WEIGHTED_ANSWER = "~%"
Const TAG_FEEDBACK = "#"

Public Rib As IRibbonUI
Sub RibbonOnLoad(ribbon As IRibbonUI)
   Set Rib = ribbon
   modRibbonI18N.InitLocale
   Rib.ActivateTab "MoodleAdaptorAddIn" ' Name of the tab to activate
End Sub
' Main method. Converts the Word document to GIFT format.
Sub ExportToGIFT(control As IRibbonControl)

    'Make sure the document is saved before continuing
    If ActiveDocument.Saved = False Then
        'MsgBox "¡Por favor, guarde este documento con marcas en formato DOC antes de continuar! " & vbCr & _
        '"Este documento se transformará en formato GIFT con el mismo nombre y extensión TXT y se perderán las marcas", _
        '        vbExclamation, "Convertidor GIFT. "
        MsgBox T("MSG_SaveDocBeforeConvert"), vbExclamation, T("TITLE_Convert2GIFT")

        Set saveDialog = Dialogs(wdDialogFileSaveAs)
        saveDialog.AddToMru = True
        saveDialog.Format = wdFormatDocument
        saveDialog.Encoding = wdFormatEncodedText

        ' Cancel pressed -> Exit
        If saveDialog.Show = 0 Then
            ' StatusBar = "Not saved"
            Exit Sub
        End If
    End If

    ' StatusBar = "Convirtiendo a formato GIFT para Moodle. Espere, por favor..."
     StatusBar = T("STATUS_ConvertingToGIFT")

    ' Before conversion, document is checked for errors
    If CheckQuestionnaire = True Then
       ' Save in UTF-8 to avoid filtering characters
        'EscapeSpanishCharacters
        'EscapeNonASCII
        ConvertToGIFT ' make the conversion
        RemoveFormatting ' remove all styles
        ' Save the document in text
        With Dialogs(wdDialogFileSaveAs)
             .Format = wdFormatText
             .Encoding = 65001
             .AllowSubstitutions = False
             .Show
        End With
    Else
        ' MsgBox "Por favor, corrija los errores antes de convertir.", vbCritical, "Error"
        MsgBox T("MSG_FixErrorsBeforeConvert"), vbCritical, T("TITLE_Error")
    End If

End Sub

' Add Multiple Choice Question to the end of the active document
Sub AddMultipleChoiceQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_MULTIPLECHOICEQ, T("TPL_MultipleChoiceQ")
End Sub
' Add Single Choice Question to the end of the active document
Sub AddSingleChoiceQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_SINGLECHOICEQ, T("TPL_SingleChoiceQ")
End Sub

' Add Matching Question to the end of the active document
Sub AddMatchingQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_MATCHINGQ, T("TPL_MatchingQ")
End Sub

' Add Numerical Question to the end of the active document
Sub AddNumericalQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_NUMERICALQ, T("TPL_NumericalQ")
End Sub


' Add Short Answer Question to the end of the active document
Sub AddShortAnswerQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_SHORTANSWERQ, T("TPL_ShortAnswerQ")
End Sub

' Add feedback
Sub AddQuestionFeedback(control As IRibbonControl)
    If Selection.Range.Style = STYLE_ANSWERWEIGHT Or _
       Selection.Range.Style = STYLE_RIGHT_ANSWER Or _
       Selection.Range.Style = STYLE_WRONG_ANSWER Then
        InsertAfterRange "Inserte la retroalimentación de su selección previa o responda aquí.", _
                         STYLE_FEEDBACK, Selection.Paragraphs(1).Range
    Else
        'MsgBox "Feedback is choice or answer specific. " & vbCr & _
        '       "Por favor, ponga el cursor en la parte derecha de la respuesta. ", vbExclamation
        MsgBox T("MSG_FeedbackCursorRightSide"), vbExclamation, T("TITLE_Convert2GIFT")
    End If
End Sub

' Add Missing Word Question
Sub AddMissingWordQ(control As IRibbonControl)
    AddParagraphOfStyle STYLE_MISSINGWORDQ, T("TPL_MissingWordQ")
End Sub

' Add a true statement of the true-false question
Sub AddTrueStatement(control As IRibbonControl)
    AddParagraphOfStyle STYLE_TRUESTATEMENT, T("TPL_TrueStatement")
End Sub

' Add a false statement of the true-false question
Sub AddFalseStatement(control As IRibbonControl)
    AddParagraphOfStyle STYLE_FALSESTATEMENT, T("TPL_FalseStatement")
End Sub
' Marks the SingleChoiceQ
Public Sub MarkSingleChoiceQ(control As IRibbonControl)
    Selection.Range.Style = STYLE_SINGLECHOICEQ
End Sub
' Marks the MultipleChoiceQ
Public Sub MarkMultipleChoiceQ(control As IRibbonControl)
        Selection.Range.Style = STYLE_MULTIPLECHOICEQ
End Sub
' Marks the right answer
Public Sub MarkTrueAnswer(control As IRibbonControl)
    If Selection.Range.Style = STYLE_WRONG_ANSWER Then
        Selection.Range.Style = STYLE_RIGHT_ANSWER
    ' ElseIf Selection.Range.Style = STYLE_RIGHT_ANSWER Then
    Else
        Selection.Range.Style = STYLE_WRONG_ANSWER
    End If
End Sub
' Marks the unmarked answer
Public Sub SetUnmarked(control As IRibbonControl)
    Selection.Range.Style = STYLE_UNMARKED_ANSWER
End Sub

' Add a new paragraph with a specified style and text
' Inserted text is selected
Private Sub AddParagraphOfStyle(aStyle, text)
    Set myRange = ActiveDocument.Content
    With myRange
        .EndOf Unit:=wdParagraph, Extend:=wdMove
        .InsertParagraphBefore
        .Move Unit:=wdParagraph, Count:=1
        .Style = aStyle
        .InsertBefore text
        .Select
    End With
End Sub



' Special Characters ~ = # { } control the operation of the Moodle's GIFT filter and
' cannot be used as a normal text within questions. However, if you want to use one
' of these characters, for example to show a mathematical formula in a question, you need
' to escape the control characters, i.e. putting a backslash (\) before a control
' character.
Private Sub EscapeControlCharacters()
    With ActiveDocument.Content.Find
        .Execute FindText:="~", ReplaceWith:="\~", _
        Format:=False, Replace:=wdReplaceAll

        .Execute FindText:="=", ReplaceWith:="\=", _
        Format:=False, Replace:=wdReplaceAll

        .Execute FindText:="#", ReplaceWith:="\#", _
        Format:=False, Replace:=wdReplaceAll

        .Execute FindText:="{", ReplaceWith:="\{", _
        Format:=False, Replace:=wdReplaceAll

        .Execute FindText:="}", ReplaceWith:="\}", _
        Format:=False, Replace:=wdReplaceAll
    End With
End Sub

' Moodle requires a dot (.) as a decimal separator. Thus, all comma separators need to
' be converted.
Private Sub ConvertDecimalSeparator(ByVal aRange As Range)
    aRange.Find.Execute FindText:=",", ReplaceWith:=".", _
    Format:=False, Replace:=wdReplaceAll
End Sub

' Remove all formatting from the document
Private Sub RemoveFormatting()
    Selection.WholeStory
    Selection.ClearFormatting
End Sub

' Count the number of paragraphs having the specified
' style in the defined range
Function CountStylesInRange(aStyle, startPoint, endPoint) As Integer
   Dim counter
   Set aRange = ActiveDocument.Range(Start:=startPoint, End:=endPoint)
   endP = aRange.End  'store end point
   counter = 0

   With aRange.Find
        .ClearFormatting
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Style = aStyle
        .Format = True
        Do While .Execute(Wrap:=wdFindStop) = True
            If aRange.End > endP Then
                Exit Do
            Else
                counter = counter + 1    ' Increment Counter.
            End If
        Loop

    End With
    CountStylesInRange = counter
End Function

' Checks every paragraph in the document and defines Moodle tags
' accordingly. Empty paragraphs are deleted as well as paragraphs
' that are specified with an unknown/illegal style.
Private Sub ConvertToGIFT()

    Dim startOfQuestion, endOfQuestion, setEndPoint

    EscapeControlCharacters ' escape all special characters

    setEndPoint = False ' indicates whether the question end point should be set

    ' Check each paragraph at a time and specify needed tags
    For Each para In ActiveDocument.Paragraphs

        ' Delete empty paragraphs
        If para.Range = vbCr Then
            para.Range.Delete ' delete all empty paragraphs
        'Question style found
        ElseIf para.Range.Style.NameLocal = STYLE_MULTIPLECHOICEQ Or _
            para.Range.Style.NameLocal = STYLE_SINGLECHOICEQ Or _
            para.Range.Style.NameLocal = STYLE_MATCHINGQ Or _
            para.Range.Style.NameLocal = STYLE_NUMERICALQ Or _
            para.Range.Style.NameLocal = STYLE_SHORTANSWERQ Then

            questionType = para.Range.Style.NameLocal ' Save the question type

            ' A new question has been found, the previous question need to be "closed"
            If setEndPoint Then InsertQuestionEndTag para.Range.Start

            ' Add start of question comments
            para.Range.InsertBefore vbCr & "// Start of question: " & questionType & vbCr

            startOfQuestion = para.Range.Start
            setEndPoint = True
            InsertAfterBeforeCR TAG_QUESTION_START, para.Range

        ' True statement found
        ElseIf para.Range.Style.NameLocal = STYLE_TRUESTATEMENT Or _
            para.Range.Style.NameLocal = STYLE_FALSESTATEMENT Or _
            para.Range.Style.NameLocal = STYLE_MISSINGWORDQ Or _
            para.Range.Style.NameLocal = STYLE_BLANK_WORD Then

            questionType = para.Range.Style.NameLocal

            If setEndPoint Then
                InsertQuestionEndTag para.Range.Start
                startOfQuestion = para.Range.Start
            End If

            para.Range.InsertBefore vbCr & "// Start of question: " & questionType & vbCr
            If questionType = STYLE_TRUESTATEMENT Then
                InsertAfterBeforeCR TAG_TRUE_CHOICE, para.Range
            ElseIf questionType = STYLE_FALSESTATEMENT Then
                InsertAfterBeforeCR TAG_FALSE_CHOICE, para.Range
            ElseIf questionType = STYLE_MISSINGWORDQ Then
                FindBlanks (para.Range)
            End If

            setEndPoint = False

        ' Wrong answer found
        ElseIf para.Range.Style.NameLocal = STYLE_WRONG_ANSWER Then
            ' Weighted answer found
            If StyleFound(STYLE_ANSWERWEIGHT, para.Range) = True Then
                InsertTextBeforeRange TAG_WEIGHTED_ANSWER, para.Range
            Else ' no answer weights are specified
                InsertTextBeforeRange TAG_WRONG_ANSWER, para.Range
            End If
        ' Right answer found
        ElseIf para.Range.Style.NameLocal = STYLE_RIGHT_ANSWER Then
            If questionType = STYLE_SINGLECHOICEQ Then
                InsertTextBeforeRange TAG_RIGHT_ANSWER, para.Range
            ' Weighted answer found
            ElseIf StyleFound(STYLE_ANSWERWEIGHT, para.Range) = True Then
                InsertTextBeforeRange TAG_WEIGHTED_ANSWER, para.Range
            ' Answer of the numerical question found
            ElseIf questionType = STYLE_NUMERICALQ Then
                InsertTextBeforeRange TAG_RIGHT_NUMERICAL_ANSWER, para.Range
            ' Answer of the multiple choice question found
            Else
                InsertTextBeforeRange TAG_RIGHT_ANSWER, para.Range
            End If
        ElseIf para.Range.Style.NameLocal = STYLE_UNMARKED_ANSWER Then
            InsertTextBeforeRange TAG_UNMARKED_ANSWER, para.Range
        ' left pair of the matching question
        ElseIf para.Range.Style.NameLocal = STYLE_LEFT_PAIR Then
            InsertTextBeforeRange TAG_RIGHT_ANSWER, para.Range
            InsertAfterBeforeCR TAG_MATCHINGQ_ARROW, para.Range
        ' right pair of the matching question
        ElseIf para.Range.Style.NameLocal = STYLE_RIGHT_PAIR Then
            ' Do nothing
            ' para.Range.Style = STYLE_NORMAL
        ' Question feedback
        ElseIf para.Range.Style.NameLocal = STYLE_FEEDBACK Then
           InsertTextBeforeRange TAG_FEEDBACK, para.Range
        Else ' Delete all undefined styles
            para.Range.Delete
        End If

        ' Check if the end of document
        If para.Range.End = ActiveDocument.Range.End Then
            ' Make sure the last line is not empty
            If para.Range = vbCr Then
                para.Range.Delete
            ElseIf setEndPoint = True Then
                InsertQuestionEndTag para.Range.End
                Exit For
            End If
        End If

    Next para



End Sub

' Check if the specified style is found in the range
Function StyleFound(aStyle, aRange As Range) As Boolean
   With aRange.Find
        .ClearFormatting
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Style = aStyle
        .Format = True
        .Execute Wrap:=wdFindStop
    End With
    'MsgBox "Style: " & aStyle & " Found: " & aRange.Find.Found
    StyleFound = aRange.Find.Found
End Function

' Removes answer weights from the selection
Public Sub RemoveAnswerWeightsFromTheSelection(control As IRibbonControl)
    With Selection.Find
        .ClearFormatting
        .Style = STYLE_ANSWERWEIGHT
        .text = ""
        .Replacement.text = ""
        .Forward = True
        .Format = True
        .Execute Replace:=wdReplaceAll
    End With
End Sub


' Checks the questionnaire.
' Returns true if everyhing is fine, otherwise false
Function CheckQuestionnaire() As Boolean

    Dim startOfQuestion, endOfQuestion, setEndPoint

    isOk = True
    setEndPoint = False ' indicates whether the question end point should be set
    startOfQuestion = 0
    questionType = ""

    ' Check each paragraph at a time and specify needed tags
    For Each para In ActiveDocument.Paragraphs

        ' Check if empty paragraph
        If para.Range = vbCr Then
            para.Range.Delete ' delete all empty paragraphs
            If questionType = "" Then questionType = para.Range.Style.NameLocal
        ElseIf para.Range.Style.NameLocal = STYLE_MULTIPLECHOICEQ Or _
               para.Range.Style.NameLocal = STYLE_SINGLECHOICEQ Or _
               para.Range.Style.NameLocal = STYLE_MATCHINGQ Or _
               para.Range.Style.NameLocal = STYLE_NUMERICALQ Or _
               para.Range.Style.NameLocal = STYLE_SHORTANSWERQ Then

            If setEndPoint Then
                endOfQuestion = para.Range.Start
                isOk = CheckQuestion(startOfQuestion, endOfQuestion)
                If isOk = False Then Exit For ' Exit if error is found
            End If

            startOfQuestion = para.Range.Start
            setEndPoint = True
            questionType = para.Range.Style.NameLocal

        ElseIf para.Range.Style.NameLocal = STYLE_TRUESTATEMENT Or _
               para.Range.Style.NameLocal = STYLE_FALSESTATEMENT Or _
               para.Range.Style.NameLocal = STYLE_MISSINGWORDQ Or _
               para.Range.Style.NameLocal = STYLE_BLANK_WORD Then

            If setEndPoint Then
                endOfQuestion = para.Range.Start
                isOk = CheckQuestion(startOfQuestion, endOfQuestion)
                If isOk = False Then Exit For ' Exit if error is found
                startOfQuestion = para.Range.Start
            End If

            questionType = para.Range.Style.NameLocal

            isOk = CheckQuestion(startOfQuestion, para.Range.End)
            If isOk = False Then Exit For ' Exit if error is found
            startOfQuestion = para.Range.End
            questionType = "NOT_KNOWN"
            setEndPoint = False
        ElseIf para.Range.Style.NameLocal = STYLE_RIGHT_ANSWER And _
               questionType = STYLE_NUMERICALQ Then
            ' Exit if error is found
            If CheckNumericAnswer(para.Range) = False Then Exit Function
        End If

        ' Check if the end of document
        If para.Range.End = ActiveDocument.Range.End And _
        startOfQuestion <> para.Range.End Then
           isOk = CheckQuestion(startOfQuestion, ActiveDocument.Range.End)
        End If

        If isOk = False Then Exit For ' Exit if error is found

    Next para
    CheckQuestionnaire = isOk
End Function

' Checks whether the chosen question is valid
' Returns true if the question is OK, otherwise
Function CheckQuestion(startPoint, endPoint) As Boolean
    Dim isOk
    Set aRange = ActiveDocument.Range(startPoint, endPoint)
    'aRange.Select
    'MsgBox "See Range for specifying question type." & questionType & vbCr & _
    '      "Start: " & startPoint & " End: " & endPoint
    isOk = True 'no errors
    If questionType = STYLE_SINGLECHOICEQ Or _
       questionType = STYLE_NUMERICALQ Then
      ' Check that there are one right anwer specified
        rightCount = CountStylesInRange(STYLE_RIGHT_ANSWER, startPoint, endPoint)

        If rightCount <> 1 Then
            aRange.Select
            MsgBox "Error, no hay una única respuesta correcta.", vbExclamation
            isOk = False
        End If
    ElseIf questionType = STYLE_MULTIPLECHOICEQ Or _
       questionType = STYLE_SHORTANSWERQ Then

        ' Check that there are right anwers specified
        rightCount = CountStylesInRange(STYLE_RIGHT_ANSWER, startPoint, endPoint)

        If rightCount = 0 Then
            aRange.Select
            MsgBox "Error, no hay una respuesta definida.", vbExclamation
            isOk = False
        End If

    ' MATCHING QUESTION
    ElseIf questionType = STYLE_MATCHINGQ Then

       ' Count the number of pairs
       rightPairCount = CountStylesInRange(STYLE_RIGHT_PAIR, startPoint, endPoint)
       leftPairCount = CountStylesInRange(STYLE_LEFT_PAIR, startPoint, endPoint)

       ' Too few pairs
       If leftPairCount < 3 Then
           aRange.Select
           MsgBox "Error, los pares no están definidos correctamente" & vbCr & _
                  "There must be atleast 3 mathching pairs.", vbExclamation, "Error!"
           isOk = False
       ' Error -> the number of left and right pairs is different or zero
       ElseIf rightPairCount <> leftPairCount Then
           aRange.Select
           MsgBox "Error, los pares no están definidos correctamente" & vbCr & _
                  "Los pares de la izquierda y derecha no son iguales.", vbExclamation, "Error!"
           isOk = False
       End If
    ElseIf questionType = STYLE_MISSINGWORDQ Then
        wordCount = CountStylesInRange(STYLE_BLANK_WORD, startPoint, endPoint)
        If wordCount <> 1 Then
            aRange.Select
            MsgBox "Con esta plantilla debe haber SOLO UN 'hueco en blanco' en cada pregunta.", vbExclamation, "Error!"
            isOk = False
        End If

    ElseIf questionType = STYLE_TRUESTATEMENT Or _
           questionType = STYLE_FALSESTATEMENT Then

    ' UNDEFINED QUESTION TYPE
    Else
        aRange.Select
        MsgBox "Tipo de pregunta indefinida:" & questionType & vbCr & _
               "Illegal question deleted.", vbExclamation, "Error!"
        aRange.Delete
    End If

    'MsgBox questionType & "=" & isOk
    CheckQuestion = isOk
End Function

' Check the numeric answer. Note, not checking all valid GIFT formats.
Function CheckNumericAnswer(aRange As Range) As Boolean

    CheckNumericAnswer = True ' By default OK

    ' Search for the error margin separator
    aRange.Find.Execute FindText:=":", Format:=False

    If aRange.Find.Found = False And IsNumeric(aRange) = False Then
            aRange.Select
            Response = MsgBox("Is this a right numerical answer?" & vbCr & _
                       "Your answer: " & aRange, vbYesNo, "Error?")
            If Response = vbNo Then CheckNumericAnswer = False
    End If
End Function

' Inserts text before the specified range. A new paragraph is inserted.
Sub InsertTextBeforeRange(ByVal text As String, ByVal aRange As Range)
    With aRange
         .Style = STYLE_NORMAL
        .InsertBefore text
        .Move Unit:=wdParagraph, Count:=1
    End With

End Sub

' Inserts text having trailing VbCr before the range
Sub InsertQuestionEndTag(endPoint As Integer)
    Set aRange = ActiveDocument.Range(endPoint - 1, endPoint)
    With aRange
        .InsertBefore vbCr & TAG_QUESTION_END
        .Style = STYLE_NORMAL
        .Move Unit:=wdParagraph, Count:=1
    End With

End Sub

' Inserts text at the end of the paragraph before the trailing VbCr
Sub InsertAfterBeforeCR(ByVal text As String, ByVal aRange As Range)
    aRange.End = aRange.End - 1 ' insert text before cr
    With aRange
        .InsertAfter text
        .Style = STYLE_NORMAL
        .Move Unit:=wdParagraph, Count:=1
    End With
End Sub

' Inserts text at the end of the paragraph
Sub InsertAfterRange(ByVal text As String, aStyle, ByVal aRange As Range)
    With aRange
        .EndOf Unit:=wdParagraph, Extend:=wdMove
        .InsertParagraphBefore
        .Move Unit:=wdParagraph, Count:=-1
        .Style = aStyle
        .InsertBefore text
        .Select
    End With
End Sub

' Set the answer weights of multiple choice questions.
Public Sub SetAnswerWeights(control As IRibbonControl) ' aStyle, startPoint, endPoint)
    Dim startPoint, endPoint, rightScore, wrongScore

    If Selection.Range.Style = STYLE_MULTIPLECHOICEQ Or _
       Selection.Range.Style = STYLE_SINGLECHOICEQ Then
        questionType = Selection.Range.Style
        startPoint = Selection.Paragraphs(1).Range.Start
        rightCount = 0
        wrongCount = 0
        Selection.MoveDown Unit:=wdParagraph, Count:=1

        Do While Selection.Range.Style = STYLE_RIGHT_ANSWER Or _
              Selection.Range.Style = STYLE_WRONG_ANSWER Or _
              Selection.Range.Style = STYLE_FEEDBACK Or _
              Selection.Range.Style = STYLE_ANSWERWEIGHT

            'Delete empty paragraphs
            If Selection.Paragraphs(1).Range = vbCr Then
                Selection.Paragraphs(1).Range.Delete ' delete all empty paragraphs
            ' Remove old answer weights
            ElseIf Selection.Range.Style = STYLE_ANSWERWEIGHT Then
                With Selection.Find
                    .ClearFormatting
                    .Style = STYLE_ANSWERWEIGHT
                    .text = ""
                    .Replacement.text = ""
                    .Forward = True
                    .Format = True
                    .Execute Replace:=wdReplaceOne
                End With
            End If

            ' Count the number of right and wrong answers
            If Selection.Range.Style = STYLE_RIGHT_ANSWER Then
                rightCount = rightCount + 1
            ElseIf Selection.Range.Style = STYLE_WRONG_ANSWER Then
                wrongCount = wrongCount + 1
            End If

           If Selection.Paragraphs(1).Range.End = ActiveDocument.Range.End Then
                endPoint = Selection.Paragraphs(1).Range.End
                Exit Do
           Else
                Selection.MoveDown Unit:=wdParagraph, Count:=1
                endPoint = Selection.Paragraphs(1).Range.Start
           End If
        Loop

        Set QuestionRange = ActiveDocument.Range(startPoint, endPoint)
        If questionType = STYLE_SINGLECHOICEQ And rightCount <> 1 Then
            QuestionRange.Select
            MsgBox "No hay una única respuesta correcta.", vbExclamation, "Error!"
        ElseIf rightCount < 1 Then
            QuestionRange.Select
            MsgBox "No hay ninguna respuesta correcta.", vbExclamation, "Error!"
        Else
            ' Calculate the right and wrong scores
            rightScore = Round(100 / rightCount, 3)
            If questionType = STYLE_SINGLECHOICEQ Then
                rightScore = "all"
            End If

            ' MODIFY the default scoring principle for wrong answers if necessary
            wrongScore = -Round(100 / wrongCount, 3)
            AddAnswerWeights QuestionRange, rightScore, wrongScore
        End If
    Else
        'MsgBox "Ponga el cursor en la pregunta " & vbCr & _
        '       "de Opción múltiple (de única o múltiple solución)", vbExclamation, "Error!"
        MsgBox T("MSG_Err_SetCursorInQuestion"), vbExclamation, T("TITLE_Error")
        ' Find the previous paragraph having the style of multiple choice question.
        With Selection.Find
            .ClearFormatting
            .text = ""
            .Style = STYLE_MULTIPLECHOICEQ
            .Forward = False
            .Format = True
            .MatchCase = False
            .Execute
        End With
    End If
End Sub

' Marks the blank word
Public Sub MarkBlankWord(control As IRibbonControl)

    Set aRange = ActiveDocument.Range(Start:=Selection.Words(1).Start, End:=Selection.Words(1).End - 1)
    If Selection.Words(1).Style = STYLE_BLANK_WORD Then
        aRange.Select
        Selection.ClearFormatting
    Else
        'RTrim(ActiveDocument.Words(1)).Style = STYLE_BLANK_WORD
        aRange.Style = STYLE_BLANK_WORD
    End If
End Sub

' Find all the
Private Sub FindBlanks(aRange As Range)
    'Set aRange = Selection.Paragraphs(1).Range


    'ActiveDocument.Content
    endPoint = aRange.End
    With aRange.Find
        .ClearFormatting
        .Style = STYLE_BLANK_WORD
        Do While .Execute(FindText:="", Forward:=True, Format:=True) = True And _
            Selection.Range.End < endPoint And ActiveDocument.Range.End <> Selection.Range.End
            With .Parent
                .InsertBefore "{="
                    .InsertAfter "}"
              '  .Move Unit:=wdParagraph, Count:=1
              .Move Unit:=wdWord, Count:=1
            End With
        Loop
    End With
End Sub

' Insert answer weights
Private Sub AddAnswerWeights(ByVal aRange As Range, rightScore, wrongScore)
    ' Check each paragraph at a time and specify needed tags
    For Each para In aRange.Paragraphs
        ' Check if empty paragraph
        If para.Range = vbCr Then
            para.Range.Delete ' delete all empty paragraphs
        ElseIf para.Range.Style = STYLE_RIGHT_ANSWER And rightScore <> "all" Then
            InsertAnswerWeight rightScore, para.Range
        ElseIf para.Range.Style = STYLE_WRONG_ANSWER Then
            InsertAnswerWeight wrongScore, para.Range
        End If
    Next para

End Sub

' Inserts text at the end of the chapter before the trailing VbCr
Private Sub InsertAnswerWeight(Score, ByVal aRange As Range)
    startPoint = aRange.Start
    scoreString = "" & Score & "%"
    aRange.InsertBefore scoreString
    Set newRange = ActiveDocument.Range(Start:=startPoint, End:=startPoint + Len(scoreString))
    newRange.Style = STYLE_ANSWERWEIGHT
    'Moodle requires that the decimal separator is dot, not comma.
    ConvertDecimalSeparator newRange

End Sub

Sub EscapeNonASCII()

    Dim d As Variant

    Dim s As String


    Set d = CreateObject("Scripting.Dictionary")
    d.Add "á", "&aacute;"
    d.Add "Á", "&Aacute;"
    d.Add "à", "&agrave;"
    d.Add "À", "&Agrave;"
    d.Add "é", "&eacute;"
    d.Add "É", "&Eacute;"
    d.Add "è", "&egrave;"
    d.Add "È", "&Egrave;"
    d.Add "í", "&iacute;"
    d.Add "Í", "&Iacute;"
    d.Add "ì", "&igrave;"
    d.Add "Ì", "&Igrave;"
    d.Add "ó", "&oacute;"
    d.Add "Ó", "&Oacute;"
    d.Add "ò", "&ograve;"
    d.Add "Ò", "&Ograve;"
    d.Add "ú", "&uacute;"
    d.Add "Ú", "&Uacute;"
    d.Add "ù", "&ugrave;"
    d.Add "Ù", "&Ugrave;"
    d.Add "ü", "&Uuml;"
    d.Add "Ü", "&Uuml;"
    d.Add "ñ", "&ntilde"
    d.Add "Ñ", "&Ntilde"
    d.Add "ß", "&szlig;"
    d.Add "ç", "&ccedil;"
    d.Add "Ç", "&Ccedil;"
    d.Add "¿", "&iquest;"
    d.Add "¡", "&iexcl;"
    d.Add "€", "&euro;"
    d.Add "£", "&pound;"
    d.Add "ª", "&ordf;"
    d.Add "²", "&sup2;"
    d.Add "³", "&sup3;"
    d.Add "´", "&acute;"
    d.Add "º", "&ordm;"
    d.Add "‘", "&lsquo;"
    d.Add "’", "&rsquo;"
    d.Add "“", "&ldquo;"
    d.Add "”", "&rdquo;"
    d.Add "…", "..."
    d.Add "~", "&sim;"
    d.Add "½", "&frac12;"
    d.Add "¼", "&frac14;"
    d.Add "¾", "&frac34;"
    d.Add "×", "&times;"
    d.Add "®", "&reg;"
    d.Add "©", "&copy;"
    d.Add "™", "&trade;"
    d.Add "¥", "&yen;"

    d.Add "^p^p^p^p", "^p"
    d.Add "^p^p^p", "^p"
    d.Add "^p^p", "^p"

    For Each key In d.keys
        s = s & key & " "
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .ClearAllFuzzyOptions
            .text = key
            .Replacement.text = d.Item(key)
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = False
            .MatchWholeWord = False
            .MatchByte = False
            .MatchAllWordForms = False
            .MatchSoundsLike = False
            .MatchFuzzy = False
            .MatchWildcards = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
    Next

End Sub

Sub EscapeSpanishCharacters()
'
' test Macro
' Macro grabada el 16/03/2012 por alfredinho
'
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "á"
        .Replacement.text = "&aacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "é"
        .Replacement.text = "&eacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "í"
        .Replacement.text = "&iacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "ó"
        .Replacement.text = "&oacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "ú"
        .Replacement.text = "&uacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "ñ"
        .Replacement.text = "&ntilde;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "¿"
        .Replacement.text = "&iquest;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "€"
        .Replacement.text = "&euro;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Á"
        .Replacement.text = "&Aacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "É"
        .Replacement.text = "&Eacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Í"
        .Replacement.text = "&Iacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Ó"
        .Replacement.text = "&Oacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll

    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Ú"
        .Replacement.text = "&Uacute;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "Ñ"
        .Replacement.text = "&Ntilde;"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = True
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
    Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll
        Selection.Find.ClearFormatting
    Selection.Find.Replacement.ClearFormatting
    With Selection.Find
        .text = "^p^p"
        .Replacement.text = "^p"
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    Selection.Find.Execute Replace:=wdReplaceAll




End Sub
Sub ShowHelp(control As IRibbonControl)
 ' MsgBox "Busca las instrucciones en la Guia de Herramientas online de la Universidad" & vbCr & COPYRIGHT, _
 '               vbExclamation, "Convertidor GIFT. "
 MsgBox T("MSG_HelpOnlineGuide") & vbCr & COPYRIGHT, _
                vbExclamation, T("TITLE_ConvertorGIFT")
End Sub
