Option Explicit

Public g_Ribbon As IRibbonUI
Private currentLocale As String

'---------------------------------------
' Inicialización del Ribbon
'---------------------------------------
Public Sub RibbonOnLoad(ribbon As IRibbonUI)
    Set g_Ribbon = ribbon
    InitLocale
End Sub
Public Sub testTranslate()
    InitLocale
    currentLocale = "fr"
    MsgBox (T("LBL_MoodleAdaptorAddIn"))
End Sub
Public Sub InitLocale()
    Dim uiLang As Long
    uiLang = Application.LanguageSettings.LanguageID(msoLanguageIDUI)

    Select Case uiLang
        Case 1033                      ' English
            currentLocale = "en"
        Case 3082, 1034                ' Spanish (Spain)
            currentLocale = "es"
        Case 1036, 2060, 3084, 4108    ' French (FR, BE, CA, CH)
            currentLocale = "fr"
        Case 1031, 3079, 4103, 5127, 2055   ' German (DE, AT, LU, LI, CH)
            currentLocale = "de"

        Case 1040, 2067                 ' Italian (IT, CH)
            currentLocale = "it"

        Case 1046, 2070                 ' Portuguese (BR, PT)
            currentLocale = "pt"
        Case Else
            currentLocale = "en"        ' fallback
    End Select
End Sub


'---------------------------------------
' Callbacks del Ribbon
'---------------------------------------
Public Sub GetLabel(control As IRibbonControl, ByRef returnedLabel)

    Select Case control.id

        ' Tab
        Case "MoodleAdaptorAddIn":           returnedLabel = T("LBL_MoodleAdaptorAddIn")

        ' Grupos
        Case "grpHelp":                     returnedLabel = T("LBL_grpHelp")
        Case "grpSingleChoice":             returnedLabel = T("LBL_grpSingleChoice")
        Case "grpMultipleChoice":           returnedLabel = T("LBL_grpMultipleChoice")
        Case "grpMatching":                 returnedLabel = T("LBL_grpMatching")
        Case "grpTrueFalse":                returnedLabel = T("LBL_grpTrueFalse")
        Case "grpShortAnswer":              returnedLabel = T("LBL_grpShortAnswer")
        Case "grpNumerical":                returnedLabel = T("LBL_grpNumerical")
        Case "grpMissingWord":              returnedLabel = T("LBL_grpMissingWord")
        Case "grpGeneral":                  returnedLabel = T("LBL_grpGeneral")

        ' Botones
        Case "btnHelp":                     returnedLabel = T("LBL_btnHelp")

        Case "btnNewSingleChoice":          returnedLabel = T("LBL_btnNewSingleChoice")
        Case "btnMarkSingleChoice":         returnedLabel = T("LBL_btnMarkSingleChoice")
        Case "btnRightWrongSingleChoice":   returnedLabel = T("LBL_btnRightWrongSingleChoice")
        Case "btnWeightsSingleChoice":      returnedLabel = T("LBL_btnWeightsSingleChoice")
        Case "btnRemoveWeightsSingleChoice": returnedLabel = T("LBL_btnRemoveWeightsSingleChoice")

        Case "btnNewMultipleChoice":        returnedLabel = T("LBL_btnNewMultipleChoice")
        Case "btnMarkMultipleChoice":       returnedLabel = T("LBL_btnMarkMultipleChoice")
        Case "btnRightWrongMultipleChoice": returnedLabel = T("LBL_btnRightWrongMultipleChoice")
        Case "btnWeightsMultipleChoice":    returnedLabel = T("LBL_btnWeightsMultipleChoice")
        Case "btnRemoveWeightsMultipleChoice": returnedLabel = T("LBL_btnRemoveWeightsMultipleChoice")

        Case "btnNewMatching":              returnedLabel = T("LBL_btnNewMatching")

        Case "btnTrueStatement":            returnedLabel = T("LBL_btnTrueStatement")
        Case "btnFalseStatement":           returnedLabel = T("LBL_btnFalseStatement")

        Case "btnNewShortAnswer":           returnedLabel = T("LBL_btnNewShortAnswer")

        Case "btnNewNumerical":             returnedLabel = T("LBL_btnNewNumerical")

        Case "btnNewMissingWord":           returnedLabel = T("LBL_btnNewMissingWord")
        Case "btnMarkBlank":                returnedLabel = T("LBL_btnMarkBlank")

        Case "btnFeedback":                 returnedLabel = T("LBL_btnFeedback")
        Case "btnExportToMoodle":           returnedLabel = T("LBL_btnExportToMoodle")

        Case Else
            returnedLabel = control.id   ' fallback para depuración
    End Select
End Sub

Public Sub GetSupertip(control As IRibbonControl, ByRef returnedSupertip)
    Select Case control.id

        Case "btnNewSingleChoice":          returnedSupertip = T("TIP_btnNewSingleChoice")
        Case "btnMarkSingleChoice":         returnedSupertip = T("TIP_btnMarkSingleChoice")
        Case "btnRightWrongSingleChoice":   returnedSupertip = T("TIP_btnRightWrongSingleChoice")
        Case "btnWeightsSingleChoice":      returnedSupertip = T("TIP_btnWeightsSingleChoice")
        Case "btnRemoveWeightsSingleChoice": returnedSupertip = T("TIP_btnRemoveWeightsSingleChoice")

        Case "btnNewMultipleChoice":        returnedSupertip = T("TIP_btnNewMultipleChoice")
        Case "btnMarkMultipleChoice":       returnedSupertip = T("TIP_btnMarkMultipleChoice")
        Case "btnRightWrongMultipleChoice": returnedSupertip = T("TIP_btnRightWrongMultipleChoice")
        Case "btnWeightsMultipleChoice":    returnedSupertip = T("TIP_btnWeightsMultipleChoice")
        Case "btnRemoveWeightsMultipleChoice": returnedSupertip = T("TIP_btnRemoveWeightsMultipleChoice")

        Case "btnNewMatching":              returnedSupertip = T("TIP_btnNewMatching")

        Case "btnTrueStatement":            returnedSupertip = T("TIP_btnTrueStatement")
        Case "btnFalseStatement":           returnedSupertip = T("TIP_btnFalseStatement")

        Case "btnNewShortAnswer":           returnedSupertip = T("TIP_btnNewShortAnswer")

        Case "btnNewNumerical":             returnedSupertip = T("TIP_btnNewNumerical")

        Case "btnNewMissingWord":           returnedSupertip = T("TIP_btnNewMissingWord")
        Case "btnMarkBlank":                returnedSupertip = T("TIP_btnMarkBlank")

        Case "btnFeedback":                 returnedSupertip = T("TIP_btnFeedback")
        Case "btnExportToMoodle":           returnedSupertip = T("TIP_btnExportToMoodle")

        Case "FileSave":                    returnedSupertip = T("TIP_FileSave")

        Case Else
            returnedSupertip = vbNullString
    End Select
End Sub

'---------------------------------------
' i18n: función T y recursos embebidos
'---------------------------------------
Public Function T(key As String) As String
    InitLocale
    Select Case currentLocale
        Case "es": T = T_es(key)
        Case "en": T = T_en(key)
        Case "fr": T = T_fr(key)
        Case "de": T = T_de(key)
        Case "it": T = T_it(key)
        Case "pt": T = T_pt(key)
        Case Else: T = ""
    End Select

    If T = "" Then
        T = T_en(key) & " Untranslated to: " & currentLocale & " #" & key & "#"
    End If
End Function


'---------------------------------------
' Textos en español
'---------------------------------------
Private Function T_es(key As String) As String

    Select Case key
        ' En T_es:
        Case "TITLE_Convert2GIFT":      T_es = "Convertidor GIFT."
        Case "TITLE_Error":             T_es = "Error"
        Case "TITLE_ErrorBang":         T_es = "Error!"
        Case "TITLE_ErrorQuestion":     T_es = "¿Error?"
        Case "STATUS_ConvertingToGIFT": T_es = "Convirtiendo a formato GIFT para Moodle. Espere, por favor..."
        ' Mensajes
        Case "MSG_SaveDocBeforeConvert"
            T_es = "Por favor, guarde este documento con marcas en formato DOC antes de continuar." & vbCr & _
                   "Este documento se transformará en formato GIFT con el mismo nombre y extensión TXT y se perderán las marcas."

        Case "MSG_FixErrorsBeforeConvert"
            T_es = "Por favor, corrija los errores antes de convertir."

        Case "MSG_FeedbackCursorRightSide"
            T_es = "La retroalimentación es específica de cada opción o respuesta." & vbCr & _
                   "Por favor, ponga el cursor en la parte derecha de la respuesta."

        Case "MSG_Err_NoSingleRightAnswer"
            T_es = "No hay una única respuesta correcta."

        Case "MSG_Err_NoAnyRightAnswer"
            T_es = "No hay ninguna respuesta correcta."

        Case "MSG_Err_NoAnswerDefined"
            T_es = "Error, no hay una respuesta definida."

        Case "MSG_Err_PairsInvalidMin3"
            T_es = "Error, los pares no están definidos correctamente." & vbCr & _
                   "Debe haber al menos 3 pares de correspondencia."

        Case "MSG_Err_PairsInvalidMismatch"
            T_es = "Error, los pares no están definidos correctamente." & vbCr & _
                   "El número de pares de la izquierda y de la derecha no coincide."

        Case "MSG_Err_OnlyOneBlankPerQuestion"
            T_es = "Con esta plantilla debe haber SOLO UN 'hueco en blanco' en cada pregunta."

        Case "MSG_Err_UndefinedQuestionTypePrefix"
            T_es = "Tipo de pregunta indefinida: "
        Case "MSG_Err_UndefinedQuestionTypeSuffix"
            T_es = vbCr & "Pregunta ilegal eliminada."

        Case "MSG_ConfirmNumericAnswerPrompt"
            T_es = "¿Es ésta una respuesta numérica correcta?" & vbCr & _
                   "Su respuesta: "

        Case "MSG_Err_SetCursorInQuestion"
            T_es = "Ponga el cursor en la pregunta " & vbCr & _
                   "de Opción múltiple (de única o múltiple solución)."

        Case "MSG_HelpOnlineGuide"
            T_es = "Busca las instrucciones en la Guía de Herramientas online de la Universidad"


        ' Labels
        Case "LBL_MoodleAdaptorAddIn":           T_es = "Exportar a Moodle"

        Case "LBL_grpHelp":                      T_es = "Ayuda"
        Case "LBL_grpSingleChoice":              T_es = "Elección única"
        Case "LBL_grpMultipleChoice":            T_es = "Elección múltiple"
        Case "LBL_grpMatching":                  T_es = "Relación"
        Case "LBL_grpTrueFalse":                 T_es = "Verdadero/Falso"
        Case "LBL_grpShortAnswer":               T_es = "Respuesta corta"
        Case "LBL_grpNumerical":                 T_es = "Numérica"
        Case "LBL_grpMissingWord":               T_es = "Palabra faltante"
        Case "LBL_grpGeneral":                   T_es = "General"

        Case "LBL_btnHelp":                      T_es = "Ayuda"

        Case "LBL_btnNewSingleChoice":           T_es = "Nueva elección única"
        Case "LBL_btnMarkSingleChoice":          T_es = "Marcar elección única"
        Case "LBL_btnRightWrongSingleChoice":    T_es = "Acierto/Fallo"
        Case "LBL_btnWeightsSingleChoice":       T_es = "Pesos"
        Case "LBL_btnRemoveWeightsSingleChoice": T_es = "Quita pesos"

        Case "LBL_btnNewMultipleChoice":         T_es = "Nueva elección múltiple"
        Case "LBL_btnMarkMultipleChoice":        T_es = "Marcar elección múltiple"
        Case "LBL_btnRightWrongMultipleChoice":  T_es = "Acierto/Fallo"
        Case "LBL_btnWeightsMultipleChoice":     T_es = "Pesos"
        Case "LBL_btnRemoveWeightsMultipleChoice": T_es = "Quita pesos"

        Case "LBL_btnNewMatching":               T_es = "Nueva de emparejamiento"

        Case "LBL_btnTrueStatement":             T_es = "Hecho verdadero"
        Case "LBL_btnFalseStatement":            T_es = "Hecho Falso"

        Case "LBL_btnNewShortAnswer":            T_es = "Nueva de respuesta corta"

        Case "LBL_btnNewNumerical":              T_es = "Nueva numérica"

        Case "LBL_btnNewMissingWord":            T_es = "Nueva palabra faltante"
        Case "LBL_btnMarkBlank":                 T_es = "Marcar hueco"

        Case "LBL_btnFeedback":                  T_es = "Realimentación"
        Case "LBL_btnExportToMoodle":            T_es = "Exportar a Moodle"

        ' Supertips
        Case "TIP_btnNewSingleChoice"
            T_es = "Añade una nueva pregunta de solución única. " & _
                   "Si su versión de Moodle no lo tiene resuelto, no olvide incluir una opción para 'No responder'."

        Case "TIP_btnMarkSingleChoice"
            T_es = "Marca el párrafo actual como pregunta de solución única. " & _
                   "Si su versión de Moodle no lo tiene resuelto, no olvide incluir una opción para 'No responder'."

        Case "TIP_btnRightWrongSingleChoice"
            T_es = "Cambia esta línea entre la opción correcta o equivocada sucesivamente."

        Case "TIP_btnWeightsSingleChoice"
            T_es = "Asigna pesos a las opciones de forma que se equilibren las respuestas al azar. Se puede cambiar manualmente."

        Case "TIP_btnRemoveWeightsSingleChoice"
            T_es = "Quita los pesos de las opciones."

        Case "TIP_btnNewMultipleChoice"
            T_es = "Añade al final de la página una nueva pregunta de elección múltiple con varias respuestas correctas."

        Case "TIP_btnMarkMultipleChoice"
            T_es = "Marca el párrafo actual como pregunta de elección múltiple con varias respuestas correctas."

        Case "TIP_btnRightWrongMultipleChoice"
            T_es = "Cambia esta línea entre la opción correcta o equivocada sucesivamente."

        Case "TIP_btnWeightsMultipleChoice"
            T_es = "Asigna pesos a las opciones de forma que se equilibren las respuestas al azar. Se puede cambiar manualmente."

        Case "TIP_btnRemoveWeightsMultipleChoice"
            T_es = "Quita los pesos de las opciones."

        Case "TIP_btnNewMatching"
            T_es = "Añade una nueva pregunta de relacionar opciones. " & _
                   "Cada línea después del enunciado se alterna entre opción izquierda y opción derecha."

        Case "TIP_btnTrueStatement"
            T_es = "Crea una pregunta de Verdadero/Falso cuya respuesta es Verdadero."

        Case "TIP_btnFalseStatement"
            T_es = "Crea una pregunta de Verdadero/Falso cuya respuesta es Falso."

        Case "TIP_btnNewShortAnswer"
            T_es = "Crea una pregunta cuya respuesta tiene que ser un texto exacto."

        Case "TIP_btnNewNumerical"
            T_es = "Crea una pregunta numérica cuya respuesta es un número. Puede usar rangos, p.ej: 34..36."

        Case "TIP_btnNewMissingWord"
            T_es = "Crea una pregunta en la que alguna parte se oculta al estudiante y éste debe introducirla según el contexto. " & _
                   "Solo se puede marcar una palabra faltante."

        Case "TIP_btnMarkBlank"
            T_es = "Indica qué parte del enunciado es la parte faltante."

        Case "TIP_btnFeedback"
            T_es = "Añade una realimentación para dar información al estudiante después de responder el cuestionario. " & _
                   "Hay que poner el cursor al final de la respuesta."

        Case "TIP_FileSave"
            T_es = "Grabe su fichero DOC con las preguntas antes de convertirlo a GIFT para no perder el trabajo de formateo."

        Case "TIP_btnExportToMoodle"
            T_es = "Este proceso transformará todo el documento en un fichero de texto plano con formato GIFT que se puede importar en Moodle."
        ' Plantillas
        Case "TPL_MultipleChoiceQ"
            T_es = "Escriba aquí el enunciado de la pregunta de elección múltiple."
        Case "TPL_SingleChoiceQ"
            T_es = "Escriba aquí el enunciado de la pregunta de elección única."
        Case "TPL_MatchingQ"
            T_es = "Escriba aquí la pregunta de relación. Después añada pares alternando izquierda/derecha."
        Case "TPL_NumericalQ"
            T_es = "Escriba aquí la pregunta numérica. Puede usar rangos (ej: 34..36)."
        Case "TPL_ShortAnswerQ"
            T_es = "Escriba aquí la pregunta de respuesta corta. La respuesta debe coincidir exactamente."
        Case "TPL_MissingWordQ"
            T_es = "Escriba aquí el texto con UNA sola palabra o frase que marcará como hueco."
        Case "TPL_TrueStatement"
            T_es = "Escriba aquí un enunciado VERDADERO para la pregunta Verdadero/Falso."
        Case "TPL_FalseStatement"
            T_es = "Escriba aquí un enunciado FALSO para la pregunta Verdadero/Falso."
        Case Else
            T_es = vbNullString
    End Select
End Function

'---------------------------------------
' Textos en inglés (traducción básica)
'---------------------------------------
Private Function T_en(key As String) As String
    Select Case key
        ' En T_en:
        Case "TITLE_Convert2GIFT":      T_en = "GIFT converter"
        Case "TITLE_Error":             T_en = "Error"
        Case "TITLE_ErrorBang":         T_en = "Error!"
        Case "TITLE_ErrorQuestion":     T_en = "Error?"
        Case "STATUS_ConvertingToGIFT": T_en = "Converting to GIFT format for Moodle. Please wait..."
        Case "MSG_SaveDocBeforeConvert"
            T_en = "Please save this document with tracked changes in DOC format before continuing." & vbCr & _
                   "This document will be converted to GIFT format with the same name and TXT extension and all tracked changes will be lost."

        Case "MSG_FixErrorsBeforeConvert"
            T_en = "Please fix the errors before converting."

        Case "MSG_FeedbackCursorRightSide"
            T_en = "Feedback is choice or answer specific." & vbCr & _
                   "Please place the cursor on the right-hand side of the answer."

        Case "MSG_Err_NoSingleRightAnswer"
            T_en = "There is not a single correct answer."

        Case "MSG_Err_NoAnyRightAnswer"
            T_en = "There is no correct answer."

        Case "MSG_Err_NoAnswerDefined"
            T_en = "Error, there is no answer defined."

        Case "MSG_Err_PairsInvalidMin3"
            T_en = "Error, pairs are not defined correctly." & vbCr & _
                   "There must be at least 3 matching pairs."

        Case "MSG_Err_PairsInvalidMismatch"
            T_en = "Error, pairs are not defined correctly." & vbCr & _
                   "The number of left and right items does not match."

        Case "MSG_Err_OnlyOneBlankPerQuestion"
            T_en = "With this template there must be ONLY ONE 'blank field' in each question."

        Case "MSG_Err_UndefinedQuestionTypePrefix"
            T_en = "Undefined question type: "
        Case "MSG_Err_UndefinedQuestionTypeSuffix"
            T_en = vbCr & "Illegal question deleted."

        Case "MSG_ConfirmNumericAnswerPrompt"
            T_en = "Is this a right numerical answer?" & vbCr & _
                   "Your answer: "

        Case "MSG_Err_SetCursorInQuestion"
            T_en = "Place the cursor on the question " & vbCr & _
                   "of multiple choice (single or multiple answer)."

        Case "MSG_HelpOnlineGuide"
            T_en = "See the instructions in the University's online Tools Guide."
        ' Labels
        Case "LBL_MoodleAdaptorAddIn":           T_en = "Export to Moodle"

        Case "LBL_grpHelp":                      T_en = "Help"
        Case "LBL_grpSingleChoice":              T_en = "Single choice"
        Case "LBL_grpMultipleChoice":            T_en = "Multiple choice"
        Case "LBL_grpMatching":                  T_en = "Matching"
        Case "LBL_grpTrueFalse":                 T_en = "True/False"
        Case "LBL_grpShortAnswer":               T_en = "Short answer"
        Case "LBL_grpNumerical":                 T_en = "Numerical"
        Case "LBL_grpMissingWord":               T_en = "Missing word"
        Case "LBL_grpGeneral":                   T_en = "General"

        Case "LBL_btnHelp":                      T_en = "Help"

        Case "LBL_btnNewSingleChoice":           T_en = "New single-choice"
        Case "LBL_btnMarkSingleChoice":          T_en = "Mark single-choice"
        Case "LBL_btnRightWrongSingleChoice":    T_en = "Right/Wrong"
        Case "LBL_btnWeightsSingleChoice":       T_en = "Weights"
        Case "LBL_btnRemoveWeightsSingleChoice": T_en = "Remove weights"

        Case "LBL_btnNewMultipleChoice":         T_en = "New multiple-choice"
        Case "LBL_btnMarkMultipleChoice":        T_en = "Mark multiple-choice"
        Case "LBL_btnRightWrongMultipleChoice":  T_en = "Right/Wrong"
        Case "LBL_btnWeightsMultipleChoice":     T_en = "Weights"
        Case "LBL_btnRemoveWeightsMultipleChoice": T_en = "Remove weights"

        Case "LBL_btnNewMatching":               T_en = "New matching"

        Case "LBL_btnTrueStatement":             T_en = "True statement"
        Case "LBL_btnFalseStatement":            T_en = "False statement"

        Case "LBL_btnNewShortAnswer":            T_en = "New short answer"

        Case "LBL_btnNewNumerical":              T_en = "New numerical"

        Case "LBL_btnNewMissingWord":            T_en = "New missing word"
        Case "LBL_btnMarkBlank":                 T_en = "Mark gap"

        Case "LBL_btnFeedback":                  T_en = "Feedback"
        Case "LBL_btnExportToMoodle":            T_en = "Export to Moodle"

        ' Supertips
        Case "TIP_btnNewSingleChoice"
            T_en = "Adds a new single-choice question. " & _
                   "If your Moodle version does not handle it, do not forget to include a 'No answer' option."

        Case "TIP_btnMarkSingleChoice"
            T_en = "Marks the current paragraph as a single-choice question. " & _
                   "If your Moodle version does not handle it, include a 'No answer' option."

        Case "TIP_btnRightWrongSingleChoice"
            T_en = "Toggles this line between correct and incorrect option."

        Case "TIP_btnWeightsSingleChoice"
            T_en = "Assigns weights to options to balance random guessing. You can edit them manually."

        Case "TIP_btnRemoveWeightsSingleChoice"
            T_en = "Removes weights from the options."

        Case "TIP_btnNewMultipleChoice"
            T_en = "Adds a new multiple-choice question with several correct answers at the end of the page."

        Case "TIP_btnMarkMultipleChoice"
            T_en = "Marks the current paragraph as a multiple-choice question with several correct answers."

        Case "TIP_btnRightWrongMultipleChoice"
            T_en = "Toggles this line between correct and incorrect option."

        Case "TIP_btnWeightsMultipleChoice"
            T_en = "Assigns weights to the options to balance random guessing. You can edit them manually."

        Case "TIP_btnRemoveWeightsMultipleChoice"
            T_en = "Removes weights from the options."

        Case "TIP_btnNewMatching"
            T_en = "Adds a new matching question. Each line after the stem alternates between left and right option."

        Case "TIP_btnTrueStatement"
            T_en = "Creates a True/False question whose correct answer is True."

        Case "TIP_btnFalseStatement"
            T_en = "Creates a True/False question whose correct answer is False."

        Case "TIP_btnNewShortAnswer"
            T_en = "Creates a question whose answer must be an exact text."

        Case "TIP_btnNewNumerical"
            T_en = "Creates a numerical question. The answer is a number and ranges are allowed, e.g. 34..36."

        Case "TIP_btnNewMissingWord"
            T_en = "Creates a missing-word question; part of the text is hidden and the student must fill it from context. Only one word can be marked per gap."

        Case "TIP_btnMarkBlank"
            T_en = "Indicates which part of the stem is the missing part."

        Case "TIP_btnFeedback"
            T_en = "Adds feedback to show to the student after answering the quiz. Place the cursor at the end of the answer."

        Case "TIP_FileSave"
            T_en = "Save your DOC file with the questions before converting to GIFT to avoid losing formatting work."

        Case "TIP_btnExportToMoodle"
            T_en = "This process converts the whole document into a plain text file in GIFT format that can be imported into Moodle."
        ' Plantillas
        Case "TPL_MultipleChoiceQ"
            T_en = "Write here the stem of the multiple-choice question."
        Case "TPL_SingleChoiceQ"
            T_en = "Write here the stem of the single-choice question."
        Case "TPL_MatchingQ"
            T_en = "Write the matching question stem. Then add pairs alternating left/right."
        Case "TPL_NumericalQ"
            T_en = "Write the numerical question here. Ranges allowed (e.g. 34..36)."
        Case "TPL_ShortAnswerQ"
            T_en = "Write the short answer question here. Answer must match exactly."
        Case "TPL_MissingWordQ"
            T_en = "Write the text with ONLY ONE word or phrase to mark as blank."
        Case "TPL_TrueStatement"
            T_en = "Write a TRUE statement for the True/False question."
        Case "TPL_FalseStatement"
            T_en = "Write a FALSE statement for the True/False question."
        Case Else
            T_en = vbNullString
    End Select
End Function

Private Function T_fr(key As String) As String
    Select Case key

        '==============================================================
        ' TÍTULOS GENERALES
        '==============================================================
        Case "TITLE_Convert2GIFT":        T_fr = "Convertisseur GIFT"
        Case "TITLE_Error":               T_fr = "Erreur"
        Case "TITLE_ErrorBang":           T_fr = "Erreur !"
        Case "TITLE_ErrorQuestion":       T_fr = "Erreur ?"

        '==============================================================
        ' ESTADO (StatusBar)
        '==============================================================
        Case "STATUS_ConvertingToGIFT"
            T_fr = "Conversion au format GIFT pour Moodle. Veuillez patienter..."

        '==============================================================
        ' MENSAJES DE Convert2GIFT.bas
        '==============================================================
        Case "MSG_SaveDocBeforeConvert"
            T_fr = "Veuillez enregistrer ce document avec le suivi des modifications au format DOC avant de continuer." & vbCr & _
                   "Ce document sera converti au format GIFT avec le même nom et une extension TXT, et toutes les marques de révision seront perdues."

        Case "MSG_FixErrorsBeforeConvert"
            T_fr = "Veuillez corriger les erreurs avant la conversion."

        Case "MSG_FeedbackCursorRightSide"
            T_fr = "La rétroaction est spécifique à chaque option ou réponse." & vbCr & _
                   "Veuillez placer le curseur sur la partie droite de la réponse."

        Case "MSG_Err_NoSingleRightAnswer"
            T_fr = "Il n’y a pas une seule réponse correcte."

        Case "MSG_Err_NoAnyRightAnswer"
            T_fr = "Il n’y a aucune réponse correcte."

        Case "MSG_Err_NoAnswerDefined"
            T_fr = "Erreur : aucune réponse n’est définie."

        Case "MSG_Err_PairsInvalidMin3"
            T_fr = "Erreur : les paires ne sont pas définies correctement." & vbCr & _
                   "Il doit y avoir au moins trois paires à faire correspondre."

        Case "MSG_Err_PairsInvalidMismatch"
            T_fr = "Erreur : les paires ne sont pas définies correctement." & vbCr & _
                   "Le nombre d’éléments à gauche et à droite n’est pas le même."

        Case "MSG_Err_OnlyOneBlankPerQuestion"
            T_fr = "Avec ce modèle, il ne peut y avoir QU’UN SEUL champ vide par question."

        Case "MSG_Err_UndefinedQuestionTypePrefix"
            T_fr = "Type de question non défini : "
        Case "MSG_Err_UndefinedQuestionTypeSuffix"
            T_fr = vbCr & "Question illégale supprimée."

        Case "MSG_ConfirmNumericAnswerPrompt"
            T_fr = "S’agit-il d’une réponse numérique correcte ?" & vbCr & _
                   "Votre réponse : "

        Case "MSG_Err_SetCursorInQuestion"
            T_fr = "Placez le curseur dans une question " & vbCr & _
                   "de type choix multiple (réponse unique ou multiple)."

        Case "MSG_HelpOnlineGuide"
            T_fr = "Consultez les instructions dans le guide en ligne des outils universitaires."

        '==============================================================
        ' ETIQUETAS DEL RIBBON (LBL_)
        '==============================================================

        ' --- Tab ---
        Case "LBL_MoodleAdaptorAddIn":           T_fr = "Exporter vers Moodle"

        ' --- Groupes ---
        Case "LBL_grpHelp":                      T_fr = "Aide"
        Case "LBL_grpSingleChoice":              T_fr = "Choix unique"
        Case "LBL_grpMultipleChoice":            T_fr = "Choix multiple"
        Case "LBL_grpMatching":                  T_fr = "Correspondance"
        Case "LBL_grpTrueFalse":                 T_fr = "Vrai/Faux"
        Case "LBL_grpShortAnswer":               T_fr = "Réponse courte"
        Case "LBL_grpNumerical":                 T_fr = "Numérique"
        Case "LBL_grpMissingWord":               T_fr = "Mot manquant"
        Case "LBL_grpGeneral":                   T_fr = "Général"

        ' --- Bouton Aide ---
        Case "LBL_btnHelp":                      T_fr = "Aide"

        ' --- Choix unique ---
        Case "LBL_btnNewSingleChoice":           T_fr = "Nouveau choix unique"
        Case "LBL_btnMarkSingleChoice":          T_fr = "Marquer choix unique"
        Case "LBL_btnRightWrongSingleChoice":    T_fr = "Correct/Incorrect"
        Case "LBL_btnWeightsSingleChoice":       T_fr = "Poids"
        Case "LBL_btnRemoveWeightsSingleChoice": T_fr = "Supprimer poids"

        ' --- Choix multiple ---
        Case "LBL_btnNewMultipleChoice":         T_fr = "Nouveau choix multiple"
        Case "LBL_btnMarkMultipleChoice":        T_fr = "Marquer choix multiple"
        Case "LBL_btnRightWrongMultipleChoice":  T_fr = "Correct/Incorrect"
        Case "LBL_btnWeightsMultipleChoice":     T_fr = "Poids"
        Case "LBL_btnRemoveWeightsMultipleChoice": T_fr = "Supprimer poids"

        ' --- Correspondance ---
        Case "LBL_btnNewMatching":               T_fr = "Nouvelle correspondance"

        ' --- Vrai/Faux ---
        Case "LBL_btnTrueStatement":             T_fr = "Énoncé vrai"
        Case "LBL_btnFalseStatement":            T_fr = "Énoncé faux"

        ' --- Réponse courte ---
        Case "LBL_btnNewShortAnswer":            T_fr = "Nouvelle réponse courte"

        ' --- Numérique ---
        Case "LBL_btnNewNumerical":              T_fr = "Nouvelle numérique"

        ' --- Mot manquant ---
        Case "LBL_btnNewMissingWord":            T_fr = "Nouveau mot manquant"
        Case "LBL_btnMarkBlank":                 T_fr = "Marquer blanc"

        ' --- Général ---
        Case "LBL_btnFeedback":                  T_fr = "Rétroaction"
        Case "LBL_btnExportToMoodle":            T_fr = "Exporter vers Moodle"

        '==============================================================
        ' SUPERTIPS DEL RIBBON (TIP_)
        '==============================================================

        Case "TIP_btnNewSingleChoice"
            T_fr = "Ajoute une nouvelle question à choix unique. " & _
                   "Si votre version de Moodle ne le gère pas, pensez à ajouter une option 'Ne répond pas'."

        Case "TIP_btnMarkSingleChoice"
            T_fr = "Marque le paragraphe actuel comme une question à choix unique."

        Case "TIP_btnRightWrongSingleChoice"
            T_fr = "Bascule cette ligne entre réponse correcte et incorrecte."

        Case "TIP_btnWeightsSingleChoice"
            T_fr = "Attribue des poids aux options pour équilibrer les réponses aléatoires."

        Case "TIP_btnRemoveWeightsSingleChoice"
            T_fr = "Supprime les poids des options sélectionnées."

        Case "TIP_btnNewMultipleChoice"
            T_fr = "Ajoute une nouvelle question à choix multiple avec plusieurs réponses correctes."

        Case "TIP_btnMarkMultipleChoice"
            T_fr = "Marque le paragraphe actuel comme une question à choix multiple."

        Case "TIP_btnRightWrongMultipleChoice"
            T_fr = "Bascule cette ligne entre réponse correcte et incorrecte."

        Case "TIP_btnWeightsMultipleChoice"
            T_fr = "Attribue des poids aux réponses pour équilibrer le hasard."

        Case "TIP_btnRemoveWeightsMultipleChoice"
            T_fr = "Supprime les poids des options sélectionnées."

        Case "TIP_btnNewMatching"
            T_fr = "Ajoute une question de correspondance. Chaque ligne alterne élément gauche / élément droit."

        Case "TIP_btnTrueStatement"
            T_fr = "Crée une question Vrai/Faux dont la réponse correcte est Vrai."

        Case "TIP_btnFalseStatement"
            T_fr = "Crée une question Vrai/Faux dont la réponse correcte est Faux."

        Case "TIP_btnNewShortAnswer"
            T_fr = "Crée une question dont la réponse doit être un texte exact."

        Case "TIP_btnNewNumerical"
            T_fr = "Crée une question numérique (peut accepter une valeur ou un intervalle, ex : 34..36)."

        Case "TIP_btnNewMissingWord"
            T_fr = "Crée une question avec un mot manquant. L’étudiant doit compléter d’après le contexte."

        Case "TIP_btnMarkBlank"
            T_fr = "Indique quel mot ou élément sera remplacé par un blanc."

        Case "TIP_btnFeedback"
            T_fr = "Ajoute une rétroaction générale pour la question. Placez le curseur à la fin de la réponse."

        Case "TIP_FileSave"
            T_fr = "Enregistrez le fichier DOC avant de convertir vers GIFT pour éviter toute perte de travail."

        Case "TIP_btnExportToMoodle"
            T_fr = "Convertit tout le document en un fichier texte au format GIFT importable dans Moodle."

        Case "TPL_MultipleChoiceQ"
            T_fr = "Saisissez ici l’énoncé de la question à choix multiple."
        Case "TPL_SingleChoiceQ"
            T_fr = "Saisissez ici l’énoncé de la question à choix unique."
        Case "TPL_MatchingQ"
            T_fr = "Saisissez l’énoncé de la correspondance puis ajoutez les paires (gauche/droite)."
        Case "TPL_NumericalQ"
            T_fr = "Saisissez ici la question numérique. Intervalles possibles (ex : 34..36)."
        Case "TPL_ShortAnswerQ"
            T_fr = "Saisissez ici la question à réponse courte. La réponse doit correspondre exactement."
        Case "TPL_MissingWordQ"
            T_fr = "Saisissez le texte avec UNE seule partie (mot ou expression) à marquer comme blanc."
        Case "TPL_TrueStatement"
            T_fr = "Saisissez un énoncé VRAI pour la question Vrai/Faux."
        Case "TPL_FalseStatement"
            T_fr = "Saisissez un énoncé FAUX pour la question Vrai/Faux."

        '==============================================================
        ' FALLBACK
        '==============================================================
        Case Else
            T_fr = vbNullString
    End Select
End Function

Private Function T_de(key As String) As String
    Select Case key
        ' Títulos
        Case "TITLE_Convert2GIFT":        T_de = "GIFT-Konverter"
        Case "TITLE_Error":               T_de = "Fehler"
        Case "TITLE_ErrorBang":           T_de = "Fehler!"
        Case "TITLE_ErrorQuestion":       T_de = "Fehler?"

        ' Estado
        Case "STATUS_ConvertingToGIFT"
            T_de = "Konvertierung in das GIFT-Format für Moodle. Bitte warten..."

        ' Mensajes
        Case "MSG_SaveDocBeforeConvert"
            T_de = "Bitte speichern Sie dieses Dokument mit Überarbeitungen im DOC-Format, bevor Sie fortfahren." & vbCr & _
                   "Dieses Dokument wird mit demselben Namen und der Erweiterung TXT in das GIFT-Format konvertiert; alle Überarbeitungen gehen verloren."

        Case "MSG_FixErrorsBeforeConvert"
            T_de = "Bitte beheben Sie die Fehler, bevor Sie konvertieren."

        Case "MSG_FeedbackCursorRightSide"
            T_de = "Die Rückmeldung bezieht sich auf die jeweilige Option oder Antwort." & vbCr & _
                   "Bitte setzen Sie den Cursor auf die rechte Seite der Antwort."

        Case "MSG_Err_NoSingleRightAnswer"
            T_de = "Es gibt nicht genau eine richtige Antwort."

        Case "MSG_Err_NoAnyRightAnswer"
            T_de = "Es gibt keine richtige Antwort."

        Case "MSG_Err_NoAnswerDefined"
            T_de = "Fehler: Es ist keine Antwort definiert."

        Case "MSG_Err_PairsInvalidMin3"
            T_de = "Fehler: Die Paare sind nicht korrekt definiert." & vbCr & _
                   "Es müssen mindestens 3 Zuordnungspaare vorhanden sein."

        Case "MSG_Err_PairsInvalidMismatch"
            T_de = "Fehler: Die Paare sind nicht korrekt definiert." & vbCr & _
                   "Die Anzahl der linken und rechten Elemente stimmt nicht überein."

        Case "MSG_Err_OnlyOneBlankPerQuestion"
            T_de = "In dieser Vorlage darf es PRO Frage nur genau EIN leeres Feld geben."

        Case "MSG_Err_UndefinedQuestionTypePrefix"
            T_de = "Undefinierter Fragetyp: "
        Case "MSG_Err_UndefinedQuestionTypeSuffix"
            T_de = vbCr & "Ungültige Frage gelöscht."

        Case "MSG_ConfirmNumericAnswerPrompt"
            T_de = "Ist dies eine korrekte numerische Antwort?" & vbCr & _
                   "Ihre Antwort: "

        Case "MSG_Err_SetCursorInQuestion"
            T_de = "Setzen Sie den Cursor in eine Frage " & vbCr & _
                   "vom Typ Mehrfachwahl (ein- oder mehrfach richtig)."

        Case "MSG_HelpOnlineGuide"
            T_de = "Siehe Anleitungen im Online-Tool-Leitfaden der Universität."

        ' Labels: Tab y grupos
        Case "LBL_MoodleAdaptorAddIn":           T_de = "Nach Moodle exportieren"

        Case "LBL_grpHelp":                      T_de = "Hilfe"
        Case "LBL_grpSingleChoice":              T_de = "Einfachauswahl"
        Case "LBL_grpMultipleChoice":            T_de = "Mehrfachauswahl"
        Case "LBL_grpMatching":                  T_de = "Zuordnung"
        Case "LBL_grpTrueFalse":                 T_de = "Wahr/Falsch"
        Case "LBL_grpShortAnswer":               T_de = "Kurzantwort"
        Case "LBL_grpNumerical":                 T_de = "Numérica"
        Case "LBL_grpMissingWord":               T_de = "Lückentext"
        Case "LBL_grpGeneral":                   T_de = "Allgemein"

        ' Labels: Botones
        Case "LBL_btnHelp":                      T_de = "Hilfe"

        Case "LBL_btnNewSingleChoice":           T_de = "Neue Einfachauswahl"
        Case "LBL_btnMarkSingleChoice":          T_de = "Einfachauswahl markieren"
        Case "LBL_btnRightWrongSingleChoice":    T_de = "Richtig/Falsch"
        Case "LBL_btnWeightsSingleChoice":       T_de = "Gewichte"
        Case "LBL_btnRemoveWeightsSingleChoice": T_de = "Gewichte entfernen"

        Case "LBL_btnNewMultipleChoice":         T_de = "Neue Mehrfachauswahl"
        Case "LBL_btnMarkMultipleChoice":        T_de = "Mehrfachauswahl markieren"
        Case "LBL_btnRightWrongMultipleChoice":  T_de = "Richtig/Falsch"
        Case "LBL_btnWeightsMultipleChoice":     T_de = "Gewichte"
        Case "LBL_btnRemoveWeightsMultipleChoice": T_de = "Gewichte entfernen"

        Case "LBL_btnNewMatching":               T_de = "Neue Zuordnung"

        Case "LBL_btnTrueStatement":             T_de = "Wahre Aussage"
        Case "LBL_btnFalseStatement":            T_de = "Falsche Aussage"

        Case "LBL_btnNewShortAnswer":            T_de = "Neue Kurzantwort"

        Case "LBL_btnNewNumerical":              T_de = "Neue numerische"

        Case "LBL_btnNewMissingWord":            T_de = "Neuer Lückentext"
        Case "LBL_btnMarkBlank":                 T_de = "Lücke markieren"

        Case "LBL_btnFeedback":                  T_de = "Rückmeldung"
        Case "LBL_btnExportToMoodle":            T_de = "Nach Moodle exportieren"

        ' Supertips
        Case "TIP_btnNewSingleChoice"
            T_de = "Fügt eine neue Einfachauswahlfrage hinzu. Falls Ihre Moodle-Version dies nicht unterstützt, fügen Sie eine Option 'Keine Antwort' hinzu."

        Case "TIP_btnMarkSingleChoice"
            T_de = "Markiert den aktuellen Absatz als Einfachauswahlfrage."

        Case "TIP_btnRightWrongSingleChoice"
            T_de = "Schaltet diese Zeile zwischen richtig und falsch um."

        Case "TIP_btnWeightsSingleChoice"
            T_de = "Weist Optionen Gewichte zu, um Raten zu balancieren. Manuelle Bearbeitung möglich."

        Case "TIP_btnRemoveWeightsSingleChoice"
            T_de = "Entfernt Gewichte von den Optionen."

        Case "TIP_btnNewMultipleChoice"
            T_de = "Fügt eine neue Mehrfachauswahlfrage mit mehreren richtigen Antworten hinzu."

        Case "TIP_btnMarkMultipleChoice"
            T_de = "Markiert den aktuellen Absatz als Mehrfachauswahlfrage."

        Case "TIP_btnRightWrongMultipleChoice"
            T_de = "Schaltet diese Zeile zwischen richtig und falsch um."

        Case "TIP_btnWeightsMultipleChoice"
            T_de = "Weist Antworten Gewichte zu, um Zufall zu minimieren. Manuelle Bearbeitung möglich."

        Case "TIP_btnRemoveWeightsMultipleChoice"
            T_de = "Entfernt Gewichte von den Optionen."

        Case "TIP_btnNewMatching"
            T_de = "Fügt eine Zuordnungsfrage hinzu. Nach dem Einleitungssatz folgt abwechselnd linkes/rechtes Element."

        Case "TIP_btnTrueStatement"
            T_de = "Erstellt eine Wahr/Falsch-Frage mit wahrer Aussage."

        Case "TIP_btnFalseStatement"
            T_de = "Erstellt eine Wahr/Falsch-Frage mit falscher Aussage."

        Case "TIP_btnNewShortAnswer"
            T_de = "Erstellt eine Frage, deren Antwort exakt übereinstimmen muss."

        Case "TIP_btnNewNumerical"
            T_de = "Erstellt eine numerische Frage. Die Antwort ist eine Zahl und Bereiche sind erlaubt, z. B. 34..36."

        Case "TIP_btnNewMissingWord"
            T_de = "Erstellt eine Lückentextfrage mit genau einer Lücke."

        Case "TIP_btnMarkBlank"
            T_de = "Gibt an, welcher Teil des Textes die Lücke ist."

        Case "TIP_btnFeedback"
            T_de = "Fügt eine Rückmeldung hinzu. Cursor ans Ende der Antwort setzen."

        Case "TIP_FileSave"
            T_de = "Speichern Sie die DOC-Datei vor der Konvertierung in GIFT, um Formatierungen nicht zu verlieren."

        Case "TIP_btnExportToMoodle"
            T_de = "Konvertiert das gesamte Dokument in eine GIFT-Textdatei, die in Moodle importiert werden kann."

        ' Plantillas
        Case "TPL_MultipleChoiceQ"
            T_de = "Geben Sie hier den Stamm der Mehrfachauswahlfrage ein."
        Case "TPL_SingleChoiceQ"
            T_de = "Geben Sie hier den Stamm der Einfachauswahlfrage ein."
        Case "TPL_MatchingQ"
            T_de = "Geben Sie die Zuordnungsfrage ein. Fügen Sie anschließend abwechselnd linke/rechte Paare hinzu."
        Case "TPL_NumericalQ"
            T_de = "Geben Sie hier die numerische Frage ein. Bereiche sind erlaubt (z. B. 34..36)."
        Case "TPL_ShortAnswerQ"
            T_de = "Geben Sie hier die Kurzantwortfrage ein. Die Antwort muss exakt übereinstimmen."
        Case "TPL_MissingWordQ"
            T_de = "Geben Sie einen Text mit GENAU EINER Lücke ein, die Sie markieren."
        Case "TPL_TrueStatement"
            T_de = "Geben Sie eine WAHRE Aussage für die Wahr/Falsch-Frage ein."
        Case "TPL_FalseStatement"
            T_de = "Geben Sie eine FALSCHE Aussage für die Wahr/Falsch-Frage ein."

        Case Else
            T_de = vbNullString
    End Select
End Function

Private Function T_it(key As String) As String
    Select Case key
        Case "TITLE_Convert2GIFT":      T_it = "Convertitore GIFT"
        Case "TITLE_Error":             T_it = "Errore"
        Case "TITLE_ErrorBang":         T_it = "Errore!"
        Case "TITLE_ErrorQuestion":     T_it = "Errore?"
        Case "STATUS_ConvertingToGIFT": T_it = "Conversione in formato GIFT per Moodle. Attendere..."
        Case "MSG_SaveDocBeforeConvert"
            T_it = "Prima di continuare, salva questo documento (con revisioni) in formato DOC." & vbCr & _
                   "Il documento sarà convertito in un file TXT in formato GIFT; le revisioni verranno perse."
        Case "MSG_FixErrorsBeforeConvert":      T_it = "Correggi gli errori prima di convertire."
        Case "MSG_FeedbackCursorRightSide":     T_it = "Il feedback è specifico di ogni opzione o risposta." & vbCr & "Posiziona il cursore a destra della risposta."
        Case "MSG_Err_NoSingleRightAnswer":     T_it = "Non c’è una sola risposta corretta."
        Case "MSG_Err_NoAnyRightAnswer":        T_it = "Non c’è alcuna risposta corretta."
        Case "MSG_Err_NoAnswerDefined":         T_it = "Errore: nessuna risposta definita."
        Case "MSG_Err_PairsInvalidMin3":        T_it = "Errore: le coppie non sono definite correttamente." & vbCr & "Devono esserci almeno 3 coppie."
        Case "MSG_Err_PairsInvalidMismatch":    T_it = "Errore: le coppie non sono definite correttamente." & vbCr & "Il numero di elementi a sinistra e a destra non coincide."
        Case "MSG_Err_OnlyOneBlankPerQuestion": T_it = "In questo modello, per ogni domanda ci deve essere SOLO UN campo vuoto."
        Case "MSG_Err_UndefinedQuestionTypePrefix": T_it = "Tipo di domanda non definito: "
        Case "MSG_Err_UndefinedQuestionTypeSuffix": T_it = vbCr & "Domanda illegale eliminata."
        Case "MSG_ConfirmNumericAnswerPrompt":  T_it = "È questa una risposta numerica corretta?" & vbCr & "La tua risposta: "
        Case "MSG_Err_SetCursorInQuestion":     T_it = "Posiziona il cursore nella domanda " & vbCr & "a scelta multipla (singola o multipla)."
        Case "MSG_HelpOnlineGuide":             T_it = "Consulta le istruzioni nella Guida online degli strumenti dell’Università."

        Case "LBL_MoodleAdaptorAddIn":          T_it = "Esporta in Moodle"
        Case "LBL_grpHelp":                     T_it = "Aiuto"
        Case "LBL_grpSingleChoice":             T_it = "Scelta singola"
        Case "LBL_grpMultipleChoice":           T_it = "Scelta multipla"
        Case "LBL_grpMatching":                 T_it = "Corrispondenza"
        Case "LBL_grpTrueFalse":                T_it = "Vero/Falso"
        Case "LBL_grpShortAnswer":              T_it = "Risposta breve"
        Case "LBL_grpNumerical":                T_it = "Numerica"
        Case "LBL_grpMissingWord":              T_it = "Parola mancante"
        Case "LBL_grpGeneral":                  T_it = "Generale"

        Case "LBL_btnHelp":                     T_it = "Aiuto"
        Case "LBL_btnNewSingleChoice":          T_it = "Nuova scelta singola"
        Case "LBL_btnMarkSingleChoice":         T_it = "Marca scelta singola"
        Case "LBL_btnRightWrongSingleChoice":   T_it = "Giusto/Sbagliato"
        Case "LBL_btnWeightsSingleChoice":      T_it = "Pesi"
        Case "LBL_btnRemoveWeightsSingleChoice": T_it = "Rimuovi pesi"

        Case "LBL_btnNewMultipleChoice":        T_it = "Nuova scelta multipla"
        Case "LBL_btnMarkMultipleChoice":       T_it = "Marca scelta multipla"
        Case "LBL_btnRightWrongMultipleChoice": T_it = "Giusto/Sbagliato"
        Case "LBL_btnWeightsMultipleChoice":    T_it = "Pesi"
        Case "LBL_btnRemoveWeightsMultipleChoice": T_it = "Rimuovi pesi"

        Case "LBL_btnNewMatching":              T_it = "Nuova corrispondenza"
        Case "LBL_btnTrueStatement":            T_it = "Affermazione vera"
        Case "LBL_btnFalseStatement":           T_it = "Affermazione falsa"
        Case "LBL_btnNewShortAnswer":           T_it = "Nuova risposta breve"
        Case "LBL_btnNewNumerical":             T_it = "Nuova numerica"
        Case "LBL_btnNewMissingWord":           T_it = "Nuova parola mancante"
        Case "LBL_btnMarkBlank":                T_it = "Marca spazio"

        Case "LBL_btnFeedback":                 T_it = "Feedback"
        Case "LBL_btnExportToMoodle":           T_it = "Esporta in Moodle"

        Case "TIP_btnNewSingleChoice":          T_it = "Aggiunge una nuova domanda a scelta singola. Se Moodle non lo gestisce, includi l’opzione 'Nessuna risposta'."
        Case "TIP_btnMarkSingleChoice":         T_it = "Marca il paragrafo corrente come domanda a scelta singola."
        Case "TIP_btnRightWrongSingleChoice":   T_it = "Commuta questa riga tra opzione corretta e errata."
        Case "TIP_btnWeightsSingleChoice":      T_it = "Assegna pesi alle opzioni per bilanciare i tentativi casuali."
        Case "TIP_btnRemoveWeightsSingleChoice": T_it = "Rimuove i pesi dalle opzioni."
        Case "TIP_btnNewMultipleChoice":        T_it = "Aggiunge una nuova domanda a scelta multipla con più risposte corrette."
        Case "TIP_btnMarkMultipleChoice":       T_it = "Marca il paragrafo corrente come scelta multipla."
        Case "TIP_btnRightWrongMultipleChoice": T_it = "Commuta questa riga tra opzione corretta e errata."
        Case "TIP_btnWeightsMultipleChoice":    T_it = "Assegna pesi alle risposte per ridurre il caso."
        Case "TIP_btnRemoveWeightsMultipleChoice": T_it = "Rimuove i pesi dalle opzioni."
        Case "TIP_btnNewMatching":              T_it = "Aggiunge una domanda di corrispondenza; le righe alternano elemento sinistro/destro."
        Case "TIP_btnTrueStatement":            T_it = "Crea una domanda Vero/Falso con affermazione vera."
        Case "TIP_btnFalseStatement":           T_it = "Crea una domanda Vero/Falso con affermazione falsa."
        Case "TIP_btnNewShortAnswer":           T_it = "Crea una domanda con risposta testuale esatta."
        Case "TIP_btnNewNumerical":             T_it = "Crea una domanda numerica; sono ammessi intervalli (es. 34..36)."
        Case "TIP_btnNewMissingWord":           T_it = "Crea una domanda con una sola parola/frase nascosta."
        Case "TIP_btnMarkBlank":                T_it = "Indica quale parte del testo è la lacuna."
        Case "TIP_btnFeedback":                 T_it = "Aggiunge un feedback; posiziona il cursore a fine risposta."
        Case "TIP_FileSave":                    T_it = "Salva il file DOC prima di convertire in GIFT per non perdere la formattazione."
        Case "TIP_btnExportToMoodle":           T_it = "Converte il documento in un file di testo GIFT importabile in Moodle."

        Case "TPL_MultipleChoiceQ":             T_it = "Scrivi qui il testo della domanda a scelta multipla."
        Case "TPL_SingleChoiceQ":               T_it = "Scrivi qui il testo della domanda a scelta singola."
        Case "TPL_MatchingQ":                   T_it = "Scrivi la domanda di corrispondenza; poi aggiungi coppie alternando sinistra/destra."
        Case "TPL_NumericalQ":                  T_it = "Scrivi qui la domanda numerica. Ammessi intervalli (es. 34..36)."
        Case "TPL_ShortAnswerQ":                T_it = "Scrivi qui la domanda a risposta breve. La risposta deve coincidere esattamente."
        Case "TPL_MissingWordQ":                T_it = "Scrivi un testo con UNA sola lacuna da marcare."
        Case "TPL_TrueStatement":               T_it = "Scrivi un’affermazione VERA per la domanda Vero/Falso."
        Case "TPL_FalseStatement":              T_it = "Scrivi un’affermazione FALSA per la domanda Vero/Falso."

        Case Else: T_it = vbNullString
    End Select
End Function

Private Function T_pt(key As String) As String
    Select Case key
        Case "TITLE_Convert2GIFT":      T_pt = "Conversor GIFT"
        Case "TITLE_Error":             T_pt = "Erro"
        Case "TITLE_ErrorBang":         T_pt = "Erro!"
        Case "TITLE_ErrorQuestion":     T_pt = "Erro?"
        Case "STATUS_ConvertingToGIFT": T_pt = "Convertendo para o formato GIFT do Moodle. Aguarde..."
        Case "MSG_SaveDocBeforeConvert"
            T_pt = "Antes de continuar, salve este documento (com revisões) em formato DOC." & vbCr & _
                   "O documento será convertido para um arquivo TXT em formato GIFT; as revisões serão perdidas."
        Case "MSG_FixErrorsBeforeConvert":      T_pt = "Corrija os erros antes de converter."
        Case "MSG_FeedbackCursorRightSide":     T_pt = "O feedback é específico para cada opção ou resposta." & vbCr & "Coloque o cursor à direita da resposta."
        Case "MSG_Err_NoSingleRightAnswer":     T_pt = "Não há uma única resposta correta."
        Case "MSG_Err_NoAnyRightAnswer":        T_pt = "Não há nenhuma resposta correta."
        Case "MSG_Err_NoAnswerDefined":         T_pt = "Erro: nenhuma resposta definida."
        Case "MSG_Err_PairsInvalidMin3":        T_pt = "Erro: os pares não estão definidos corretamente." & vbCr & "Devem existir pelo menos 3 pares."
        Case "MSG_Err_PairsInvalidMismatch":    T_pt = "Erro: os pares não estão definidos corretamente." & vbCr & "A contagem de itens à esquerda e à direita não coincide."
        Case "MSG_Err_OnlyOneBlankPerQuestion": T_pt = "Neste modelo, deve haver APENAS UM campo em branco por questão."
        Case "MSG_Err_UndefinedQuestionTypePrefix": T_pt = "Tipo de pergunta não definido: "
        Case "MSG_Err_UndefinedQuestionTypeSuffix": T_pt = vbCr & "Pergunta inválida removida."
        Case "MSG_ConfirmNumericAnswerPrompt":  T_pt = "Esta é uma resposta numérica correta?" & vbCr & "Sua resposta: "
        Case "MSG_Err_SetCursorInQuestion":     T_pt = "Coloque o cursor na questão " & vbCr & "de múltipla escolha (única ou múltipla)."
        Case "MSG_HelpOnlineGuide":             T_pt = "Veja as instruções no Guia de Ferramentas online da Universidade."

        Case "LBL_MoodleAdaptorAddIn":          T_pt = "Exportar para Moodle"
        Case "LBL_grpHelp":                     T_pt = "Ajuda"
        Case "LBL_grpSingleChoice":             T_pt = "Escolha única"
        Case "LBL_grpMultipleChoice":           T_pt = "Escolha múltipla"
        Case "LBL_grpMatching":                 T_pt = "Correspondência"
        Case "LBL_grpTrueFalse":                T_pt = "Verdadeiro/Falso"
        Case "LBL_grpShortAnswer":              T_pt = "Resposta curta"
        Case "LBL_grpNumerical":                T_pt = "Numérica"
        Case "LBL_grpMissingWord":              T_pt = "Palavra ausente"
        Case "LBL_grpGeneral":                  T_pt = "Geral"

        Case "LBL_btnHelp":                     T_pt = "Ajuda"
        Case "LBL_btnNewSingleChoice":          T_pt = "Nova escolha única"
        Case "LBL_btnMarkSingleChoice":         T_pt = "Marcar escolha única"
        Case "LBL_btnRightWrongSingleChoice":   T_pt = "Certo/Errado"
        Case "LBL_btnWeightsSingleChoice":      T_pt = "Pesos"
        Case "LBL_btnRemoveWeightsSingleChoice": T_pt = "Remover pesos"

        Case "LBL_btnNewMultipleChoice":        T_pt = "Nova escolha múltipla"
        Case "LBL_btnMarkMultipleChoice":       T_pt = "Marcar escolha múltipla"
        Case "LBL_btnRightWrongMultipleChoice": T_pt = "Certo/Errado"
        Case "LBL_btnWeightsMultipleChoice":    T_pt = "Pesos"
        Case "LBL_btnRemoveWeightsMultipleChoice": T_pt = "Remover pesos"

        Case "LBL_btnNewMatching":              T_pt = "Nova correspondência"
        Case "LBL_btnTrueStatement":            T_pt = "Afirmação verdadeira"
        Case "LBL_btnFalseStatement":           T_pt = "Afirmação falsa"
        Case "LBL_btnNewShortAnswer":           T_pt = "Nova resposta curta"
        Case "LBL_btnNewNumerical":             T_pt = "Nova numérica"
        Case "LBL_btnNewMissingWord":           T_pt = "Nova palavra ausente"
        Case "LBL_btnMarkBlank":                T_pt = "Marcar espaço"

        Case "LBL_btnFeedback":                 T_pt = "Feedback"
        Case "LBL_btnExportToMoodle":           T_pt = "Exportar para Moodle"

        Case "TIP_btnNewSingleChoice":          T_pt = "Adiciona uma nova questão de escolha única. Se o Moodle não suportar, inclua 'Sem resposta'."
        Case "TIP_btnMarkSingleChoice":         T_pt = "Marca o parágrafo atual como escolha única."
        Case "TIP_btnRightWrongSingleChoice":   T_pt = "Alterna entre opção correta e incorreta."
        Case "TIP_btnWeightsSingleChoice":      T_pt = "Atribui pesos para equilibrar palpites aleatórios."
        Case "TIP_btnRemoveWeightsSingleChoice": T_pt = "Remove os pesos das opções."
        Case "TIP_btnNewMultipleChoice":        T_pt = "Adiciona uma questão de escolha múltipla com várias respostas corretas."
        Case "TIP_btnMarkMultipleChoice":       T_pt = "Marca o parágrafo atual como escolha múltipla."
        Case "TIP_btnRightWrongMultipleChoice": T_pt = "Alterna entre opção correta e incorreta."
        Case "TIP_btnWeightsMultipleChoice":    T_pt = "Atribui pesos às respostas para reduzir o acaso."
        Case "TIP_btnRemoveWeightsMultipleChoice": T_pt = "Remove os pesos das opções."
        Case "TIP_btnNewMatching":              T_pt = "Adiciona uma questão de correspondência; linhas alternam item esquerdo/direito."
        Case "TIP_btnTrueStatement":            T_pt = "Cria uma questão Verdadeiro/Falso com afirmação verdadeira."
        Case "TIP_btnFalseStatement":           T_pt = "Cria uma questão Verdadeiro/Falso com afirmação falsa."
        Case "TIP_btnNewShortAnswer":           T_pt = "Cria uma questão cuja resposta deve ser texto exato."
        Case "TIP_btnNewNumerical":             T_pt = "Cria uma questão numérica; intervalos são permitidos (ex.: 34..36)."
        Case "TIP_btnNewMissingWord":           T_pt = "Cria uma questão com apenas uma parte oculta do texto."
        Case "TIP_btnMarkBlank":                T_pt = "Indica qual parte do texto é o espaço em branco."
        Case "TIP_btnFeedback":                 T_pt = "Adiciona feedback; coloque o cursor no fim da resposta."
        Case "TIP_FileSave":                    T_pt = "Salve o DOC antes de converter para GIFT para não perder a formatação."
        Case "TIP_btnExportToMoodle":           T_pt = "Converte o documento em um arquivo de texto GIFT importável no Moodle."

        Case "TPL_MultipleChoiceQ":             T_pt = "Escreva aqui o enunciado da questão de escolha múltipla."
        Case "TPL_SingleChoiceQ":               T_pt = "Escreva aqui o enunciado da questão de escolha única."
        Case "TPL_MatchingQ":                   T_pt = "Escreva a questão de correspondência; depois adicione pares alternando esquerda/direita."
        Case "TPL_NumericalQ":                  T_pt = "Escreva aqui a questão numérica. Intervalos permitidos (ex.: 34..36)."
        Case "TPL_ShortAnswerQ":                T_pt = "Escreva aqui a questão de resposta curta. A resposta deve coincidir exatamente."
        Case "TPL_MissingWordQ":                T_pt = "Escreva um texto com APENAS UMA lacuna para marcar."
        Case "TPL_TrueStatement":               T_pt = "Escreva uma afirmação VERDADEIRA para a questão Verdadeiro/Falso."
        Case "TPL_FalseStatement":              T_pt = "Escreva uma afirmação FALSA para a questão Verdadeiro/Falso."

        Case Else: T_pt = vbNullString
    End Select
End Function