'------------------------------------------------------------------------------
' Modul: owui.bas
' Zweck: VBA Makros, um aus Word API Calls zu einer OpenWebUI- und Ollama-Instanz zu machen
' Autor: Jonas Achermann, Kriminalgericht Luzern, jonas.achermann@lu.ch
' Letztes grosses Update: 29.06.2025 (Standalone Version, ausser Konfiguration, kein Verweis auf anderes Modul)
'------------------------------------------------------------------------------

'=== Grundkonfiguration ===
' Es Muss ein Modul owuiconfig bestehen, in welchem die Variablen für
'- OWUI_API_URL
'- OWUI_API_TOKEN
'- OWUI_MODEL
'definiert werden

' ==============================================================================
' API-Deklarationen für den Zugriff auf die Windows-Zwischenablage
' Kompatibel mit 32-Bit und 64-Bit Office-Versionen
' ==============================================================================

' Bedingte Kompilierung für moderne (VBA7) und ältere VBA-Versionen
#If VBA7 Then
    ' Für Office 2010 und neuer (VBA Version 7)

    ' Öffnet und schliesst die Zwischenablage
    Public Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    
    ' Ruft Daten aus der Zwischenablage ab
    Public Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As LongPtr
    
    ' Sperrt und entsperrt einen globalen Speicherblock
    Public Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Public Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    
    ' Kopiert einen String und ermittelt seine Länge (ANSI-Version für VBA)
    Public Declare PtrSafe Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As LongPtr) As Long
    Public Declare PtrSafe Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As LongPtr) As Long
    
    ' Windows-API-Deklaration, um ein Fenster in den Vordergrund zu bringen
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr

#Else
    ' Für Office 2007 und älter (VBA Version 6)
    
    Public Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function CloseClipboard Lib "user32" () As Long
    Public Declare Function GetClipboardData Lib "user32" (ByVal uFormat As Long) As Long
    Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Long) As Long
    Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
    
    ' Windows-API-Deklaration, um ein Fenster in den Vordergrund zu bringen
    Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
    Public Declare Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    
#End If

' Konstante für das Textformat der Zwischenablage
Public Const CF_TEXT As Long = 1


'=== Makros ===

'------------------------------------------------------------------------------
' Sub: ShowPromptForm
' Purpose: Öffnet das Formular, um einen Prompt auszuwählen
'------------------------------------------------------------------------------
Public Sub ShowPromptForm()
    frmPromptTool.Show vbModeless
End Sub

'------------------------------------------------------------------------------
' Sub: aiKonvertiereWeiblich
' Purpose: Konvertiert eine Aussage einer weiblichen Person in indirekte Rede
'------------------------------------------------------------------------------
Public Sub aiKonvertiereWeiblich()
    Dim prompt As String
    Dim result As String
    Dim commandName As String
    Dim prompt_muster As String ' Variable deklarieren
    
    commandName = "/konvertierung-weiblich" ' Beispiel-Command

    ' Prompt-Muster bei OWUI Prompts holen
    prompt_muster = GetPromptByCommandName(commandName)
    
    ' Füllvariablen in Prompt ausfüllen (bspw. Zwischenablage in {{CLIPBOARD}} einfüllen)
    prompt = InjectPrompt(prompt_muster)

    ' Ohne Streaming, weil dann Belegstelle hinzufügen nicht funktioniert
    result = StreamOWUIToWordWithModel(prompt, OWUI_MODEL, StreamToDocument:=False)

    ' Text mit Belegstelle einfügen
    Call InsertTextIntoDocument(result, mitBelegstelle:=True)
End Sub


'------------------------------------------------------------------------------
' Sub: aiKonvertiereMaennlich
' Purpose: Konvertiert eine Aussage einer weiblichen Person in indirekte Rede
'------------------------------------------------------------------------------
Public Sub aiKonvertiereMaennlich()
    Dim prompt As String
    Dim result As String
    Dim commandName As String
    Dim prompt_muster As String ' Variable deklarieren
    
    commandName = "/konvertierung-mannlich" ' Beispiel-Command

    ' Prompt-Muster bei OWUI Prompts holen
    prompt_muster = GetPromptByCommandName(commandName)
    
    ' Füllvariablen in Prompt ausfüllen (bspw. Zwischenablage in {{CLIPBOARD}} einfüllen)
    prompt = InjectPrompt(prompt_muster)

    ' Ohne Streaming, weil dann Belegstelle hinzufügen nicht funktioniert
    result = StreamOWUIToWordWithModel(prompt, OWUI_MODEL, StreamToDocument:=False)

    ' Text mit Belegstelle einfügen
    Call InsertTextIntoDocument(result, mitBelegstelle:=True)
End Sub

' --- NEU: Demo-Makro zum Testen der neuen Funktion ---
'------------------------------------------------------------------------------
' Sub: DemoPromptAbrufen
' Purpose: Demonstrates how to fetch a predefined prompt from the API
'------------------------------------------------------------------------------
Public Sub DemoPromptAbrufen()
    Dim commandName As String
    commandName = "/konvertierung-weiblich" ' Beispiel-Command

    Dim promptContent As String
    promptContent = GetPromptByCommandName(commandName)
    
    promptContent = InjectPrompt(promptContent)

    If promptContent <> "" Then
        MsgBox "Abgerufener Prompt-Inhalt für '" & commandName & "':" & vbCrLf & vbCrLf & promptContent
        InsertTextIntoDocument (promptContent)
    Else
        MsgBox "Konnte den Prompt für den Befehl '" & commandName & "' nicht finden oder abrufen.", vbExclamation
    End If
End Sub


'=== Main Inference Functions ===

'------------------------------------------------------------------------------
' Function: StreamOWUIToWordWithModel
' Purpose: Streams text from OWUI or returns it as a whole string.
' Parameters:
'   - prompt: The prompt to send to the model
'   - modelName: The name of the model to use
'   - StreamToDocument (Optional, Default=True): If True, streams response to Word.
'     If False, only returns the complete text without modifying the document.
' Returns: Complete response text
'------------------------------------------------------------------------------
Public Function StreamOWUIToWordWithModel(prompt As String, modelName As String, Optional ByVal StreamToDocument As Boolean = True) As String ' <<< NEUER PARAMETER
    On Error GoTo ErrorHandler

    Const OLLAMA_API_URL As String = "https://aiacoder-v001.kt.lunet.ch/ollama/api/generate"
    Const API_KEY As String = OWUI_API_TOKEN

    Dim httpRequest As Object
    Dim jsonData As String
    Dim responseText As String
    Dim line As String
    Dim responseValue As String

    Set httpRequest = CreateObject("MSXML2.XMLHTTP")

    jsonData = "{""model"": """ & modelName & """, ""prompt"": """ & EscapeJson(TrimEndCustom(prompt)) & """, ""stream"": true}"

    httpRequest.Open "POST", OLLAMA_API_URL, False
    httpRequest.setRequestHeader "Authorization", "Bearer " & API_KEY
    httpRequest.setRequestHeader "Content-Type", "application/json"
    httpRequest.send jsonData

    If httpRequest.Status <> 200 Then
        MsgBox "Fehler bei der HTTP-Anfrage: " & httpRequest.Status & " - " & httpRequest.StatusText
        StreamOWUIToWordWithModel = "Fehler bei der Anfrage"
        Exit Function
    End If

    responseText = ""
    Dim lines() As String
    lines = Split(httpRequest.responseText, vbLf)

    Dim i As Integer
    For i = 0 To UBound(lines)
        line = Trim(lines(i))
        If line <> "" Then
            responseValue = ExtractJsonValue(line, "response")
            If responseValue <> "" Then

                ' Nur in das Dokument schreiben, wenn der Parameter es erlaubt.
                If StreamToDocument Then
                    Selection.TypeText text:=responseValue
                    DoEvents ' DoEvents ist nur beim Streamen in die UI sinnvoll
                End If
                
                ' Der Text wird immer für die Rückgabe der Funktion zusammengesetzt.
                responseText = responseText & responseValue
            End If
        End If
    Next i

    StreamOWUIToWordWithModel = responseText
    Exit Function

ErrorHandler:
    MsgBox "Ein Fehler ist in StreamOWUIToWordWithModel aufgetreten: " & Err.Description
    StreamOWUIToWordWithModel = "Fehler"
End Function


' -- Inference Helper Functions --

'------------------------------------------------------------------------------
' Function: EscapeJson
' Purpose: Escapes special characters in a string for JSON
' Parameters:
'   - text: The text to escape
' Returns: Escaped text
'------------------------------------------------------------------------------
Private Function EscapeJson(text As String) As String
    On Error GoTo ErrorHandler

    Dim tmp As String
    tmp = text

    tmp = Replace(tmp, "\", "\\") ' Backslashes (\) -> double backslashes (\\)
    tmp = Replace(tmp, """", "\""") ' Double quotes (") -> \"
    tmp = Replace(tmp, vbCrLf, "\n") ' CRLF -> \n
    tmp = Replace(tmp, vbLf, "\n") ' LF -> \n
    tmp = Replace(tmp, vbCr, "\n") ' CR -> \n
    tmp = Replace(tmp, vbTab, "\t") ' Tab -> \t
    tmp = Replace(tmp, "/", "\/") ' Slash -> \/ (optional)

    ' Remove all \n, \t and spaces at the end
    Do While Right(tmp, 2) = "\n" Or Right(tmp, 2) = "\t" Or Right(tmp, 1) = " "
        If Right(tmp, 2) = "\n" Then
            tmp = Left(tmp, Len(tmp) - 2)
        ElseIf Right(tmp, 2) = "\t" Then
            tmp = Left(tmp, Len(tmp) - 2)
        ElseIf Right(tmp, 1) = " " Then
            tmp = Left(tmp, Len(tmp) - 1)
        End If
    Loop

    'Debug.Print "Text after Escape-Function: " & tmp

    EscapeJson = tmp
    Exit Function

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in EscapeJson"
    EscapeJson = text
End Function

'------------------------------------------------------------------------------
' Function: ExtractJsonValue
' Purpose: Extracts a value from a JSON string without a JSON parser
' Parameters:
'   - jsonStr: The JSON string
'   - key: The key to extract
' Returns: The extracted value
'------------------------------------------------------------------------------
Private Function ExtractJsonValue(jsonStr As String, key As String) As String
    On Error GoTo ErrorHandler

    Dim keyPattern As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String

    ' Create pattern for the key
    keyPattern = """" & key & """:"

    ' Find key position
    startPos = InStr(1, jsonStr, keyPattern)

    ' If key not found, return empty string
    If startPos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    ' Calculate start position for the value
    startPos = startPos + Len(keyPattern)

    ' Skip whitespace
    Do While Mid(jsonStr, startPos, 1) = " "
        startPos = startPos + 1
    Loop

    ' Check if value is a string
    If Mid(jsonStr, startPos, 1) = """" Then
        ' For strings: Find closing quote
        startPos = startPos + 1
        endPos = InStr(startPos, jsonStr, """")

        ' If no closing quote found, take rest of string
        If endPos = 0 Then
            value = Mid(jsonStr, startPos)
        Else
            value = Mid(jsonStr, startPos, endPos - startPos)
        End If
    Else
        ' For non-strings: Find next comma or }
        endPos = InStr(startPos, jsonStr, ",")
        Dim endBracePos As Long
        endBracePos = InStr(startPos, jsonStr, "}")

        If endPos = 0 Or (endBracePos > 0 And endBracePos < endPos) Then
            endPos = endBracePos
        End If

        If endPos = 0 Then
            value = Mid(jsonStr, startPos)
        Else
            value = Mid(jsonStr, startPos, endPos - startPos)
        End If
    End If

    ' Replace escape sequences
    value = Replace(value, "\""", """")
    value = Replace(value, "\\", "\")
    value = Replace(value, "\n", vbNewLine)

    ExtractJsonValue = value
    Exit Function

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in ExtractJsonValue"
    ExtractJsonValue = ""
End Function

'------------------------------------------------------------------------------
' Function: TrimEndCustom
' Purpose: Removes whitespace and control characters from the end of a string
' Parameters:
'   - text: The text to trim
' Returns: Trimmed text
'------------------------------------------------------------------------------
Private Function TrimEndCustom(text As String) As String
    On Error GoTo ErrorHandler

    ' Remove all whitespace, tabs, and control characters at the end
    Do While Len(text) > 0 And (Right(text, 1) = vbCr Or Right(text, 1) = vbLf Or _
                        Right(text, 1) = vbTab Or Right(text, 1) = " " Or _
                        AscW(Right(text, 1)) < 32)
        text = Left(text, Len(text) - 1)
    Loop

    TrimEndCustom = text
    Exit Function

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description & " in TrimEndCustom"
    TrimEndCustom = text
End Function



'=== Prompt Handling ===

'------------------------------------------------------------------------------
' Funktion: GetPromptByCommandName
' Zweck: Ruft einen auf OWUI vordefinierten Prompt via OWUI API anhand seines Befehlsnamens ab.
'        Die Funktion enthält die gesamte Logik, inklusive des JSON-Parsings.
' Parameter:
'     - commandName: Der Bezeichner des Befehls (z. B. "/konvertierung-weiblich")
' Rückgabewert: Der Inhalt des gefundenen Prompts als String. Gibt bei Fehlern oder
'              wenn nichts gefunden wird, einen leeren String zurück.
'------------------------------------------------------------------------------
Public Function GetPromptByCommandName(ByVal commandName As String) As String
    On Error GoTo ErrorHandler

    Dim request As Object
    Dim jsonResponse As String
    Dim promptContent As String
    
    ' --- Cache-Buster hinzufügen, damit MSXML2.XMLHTTP nicht den Cache verwendet ---
    Dim cacheBuster As String
    cacheBuster = "cache=" & Round(Timer * 100, 0) ' Erzeugt eine quasi-zufällige Zahl

    ' 1. HTTP-Anfrage erstellen
    Set request = CreateObject("MSXML2.XMLHTTP")
    
    ' --- Cache-Buster an die URL anhängen ---
    request.Open "GET", OWUI_API_URL & "/api/v1/prompts/list" & "?" & cacheBuster, False
    request.setRequestHeader "Authorization", "Bearer " & OWUI_API_TOKEN
    request.send

    ' 2. Antwort des Servers prüfen
    If request.Status = 200 Then
        jsonResponse = request.responseText
        'Debug.Print jsonResponse
        
        ' 3. JSON parsen und den richtigen Inhalt extrahieren (integrierte Logik)
        '--------------------------------------------------------------------------
        Dim commandMarker As String
        Dim contentMarker As String
        Dim commandPos As Long
        Dim contentStartPos As Long
        Dim i As Long
        Dim result As String
        Dim currentChar As String
        Dim nextChar As String
        
        ' Marker, nach denen wir im JSON-String suchen
        commandMarker = """command"":""" & commandName & """"
        contentMarker = """content"":"""

        ' 3.1. Finde den Start des richtigen Prompt-Objekts im JSON
        commandPos = InStr(1, jsonResponse, commandMarker, vbTextCompare)
        
        If commandPos > 0 Then
            ' 3.2. Finde den Anfang des "content"-Wertes nach dem gefundenen Befehl
            contentStartPos = InStr(commandPos, jsonResponse, contentMarker, vbTextCompare)
            
            If contentStartPos > 0 Then
                ' Position auf das erste Zeichen *nach* dem "content":"-Marker setzen
                contentStartPos = contentStartPos + Len(contentMarker)
                
                ' 3.3. JSON-String Zeichen für Zeichen durchgehen, um den Inhalt zu extrahieren
                '      und Escape-Sequenzen korrekt zu behandeln.
                result = ""
                i = contentStartPos
                Do While i <= Len(jsonResponse)
                    currentChar = Mid(jsonResponse, i, 1)
                    
                    If currentChar = "\" Then ' Escape-Sequenz gefunden
                        nextChar = Mid(jsonResponse, i + 1, 1)
                        Select Case nextChar
                            Case """"
                                result = result & """"
                            Case "n"
                                result = result & vbCrLf
                            Case "t"
                                result = result & vbTab
                            Case "\"
                                result = result & "\"
                            Case "/"
                                result = result & "/"
                            ' Fügen Sie hier bei Bedarf weitere Escape-Sequenzen hinzu
                            Case Else
                                result = result & nextChar ' Fallback: Das Zeichen einfach übernehmen
                        End Select
                        i = i + 2 ' Zwei Positionen überspringen (\ und das folgende Zeichen)
                    ElseIf currentChar = """" Then ' Potenzielles Ende des Inhalts
                        ' Prüfen, ob dies das schließende Anführungszeichen ist, indem wir
                        ' auf ein Komma oder eine geschweifte Klammer danach achten.
                        Dim nextRelevantChar As String
                        nextRelevantChar = Trim(Mid(jsonResponse, i + 1, 2))
                        
                        If Left(nextRelevantChar, 1) = "," Or Left(nextRelevantChar, 1) = "}" Then
                            Exit Do ' Das Ende des Inhalts wurde erreicht
                        Else
                            result = result & currentChar ' Es war ein Anführungszeichen im Text
                            i = i + 1
                        End If
                    Else
                        result = result & currentChar
                        i = i + 1
                    End If
                Loop
                promptContent = result
            Else
                promptContent = "" ' "content"-Feld nach dem Befehl nicht gefunden
            End If
        Else
            promptContent = "" ' Befehl nicht gefunden
        End If
        '--- Ende des integrierten Parsings ---
        
    Else
        MsgBox "Fehler beim Abrufen der Prompts." & vbCrLf & _
               "Status: " & request.Status & " - " & request.StatusText, vbCritical
        promptContent = ""
    End If

    GetPromptByCommandName = promptContent
    Set request = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Ein unerwarteter Fehler ist in GetPromptByCommandName aufgetreten: " & Err.Description, vbCritical
    GetPromptByCommandName = ""
    Set request = Nothing
End Function


'------------------------------------------------------------------------------
' Funktion: InjectPrompt
' Zweck: Ersetzt Platzhalter in einem Prompt-Text durch:
'        - Benutzereingabe über ein Dialogfenster, optional mit dynamischem
'          Anweisungstext ({{ASKINSTRUCTION:"Ihr Text"}}).
'        - Inhalte aus der Zwischenablage ({{CLIPBOARD}})
'        - den selektierten Text ({{TEXTSELECTED}})
'        - Text vor ({{TEXTBEFORE}} oder {{TEXTBEFORE:500}}) oder
'          nach ({{TEXTAFTER}} oder {{TEXTAFTER:500}}) dem Cursor.
' Parameter:
'   - promptText: Der Prompt-String, der Platzhalter enthalten kann.
'   - removeSonderzeichen: (Optional) Wenn True, werden Sonderzeichen aus
'                          dem eingefügten Text vor der Einfügung entfernt.
' Gibt zurück: Der Prompt-Text mit den ersetzten Platzhaltern.
'------------------------------------------------------------------------------
Public Function InjectPrompt(ByVal promptText As String, Optional removeSonderzeichen As Boolean = False) As String
    On Error GoTo ErrorHandler

    Dim processedText As String
    processedText = promptText
    
    ' 1. {{MODEL}}-Platzhalter entfernen
    Dim modelPlaceholderStart As Long
    Dim modelPlaceholderEnd As Long
    Dim fullModelPlaceholder As String
    
    modelPlaceholderStart = InStr(1, processedText, "{{MODEL:", vbTextCompare)
    If modelPlaceholderStart > 0 Then
        modelPlaceholderEnd = InStr(modelPlaceholderStart, processedText, "}}", vbTextCompare)
        If modelPlaceholderEnd > 0 Then
            fullModelPlaceholder = Mid(processedText, modelPlaceholderStart, modelPlaceholderEnd - modelPlaceholderStart + 2)
            ' Ersetze den gefundenen Platzhalter durch einen leeren String
            processedText = Replace(processedText, fullModelPlaceholder, "", 1, 1, vbTextCompare)
        End If
    End If

    ' 2. {{ASKINSTRUCTION}}-Platzhalter mit dynamischem Prompt behandeln
    Dim startPos As Long, endPos As Long, promptStart As Long, promptEnd As Long
    Dim fullPlaceholder As String, customPrompt As String, userInstruction As String
    
    ' Schleife, die läuft, solange {{ASKINSTRUCTION... gefunden wird.
    startPos = InStr(1, processedText, "{{ASKINSTRUCTION", vbTextCompare)
    Do While startPos > 0
        endPos = InStr(startPos, processedText, "}}", vbTextCompare)
        If endPos = 0 Then Exit Do ' Schleife verlassen, wenn Platzhalter unvollständig ist

        ' Den kompletten Platzhalter extrahieren (z.B. {{ASKINSTRUCTION:"Fokus?"}})
        fullPlaceholder = Mid(processedText, startPos, endPos - startPos + 2)

        ' Prüfen, ob ein benutzerdefinierter Prompt vorhanden ist (im Format :"...")
        promptStart = InStr(1, fullPlaceholder, ":" & Chr(34), vbTextCompare) ' Sucht nach :"
        
        If promptStart > 0 Then
            ' Finde das schliessende Anführungszeichen
            promptEnd = InStr(promptStart + 2, fullPlaceholder, Chr(34), vbTextCompare)
            If promptEnd > 0 Then
                ' Extrahiere den Anweisungstext
                customPrompt = Mid(fullPlaceholder, promptStart + 2, promptEnd - (promptStart + 2))
            Else
                ' Fallback, falls schliessendes " fehlt
                customPrompt = "Bitte Anweisung eingeben (Formatfehler im Prompt):"
            End If
        Else
            ' Fallback: Standard-Prompt, wenn nur {{ASKINSTRUCTION}} verwendet wird
            customPrompt = "Bitte geben Sie die gewünschte Anweisung oder Information ein:"
        End If

        ' Öffne die InputBox mit dem extrahierten oder dem Standard-Prompt
        userInstruction = InputBox(customPrompt, "Zusätzliche Anweisung")

        If removeSonderzeichen Then
            userInstruction = SonderzeichenEntfernen(userInstruction)
        End If
        
        ' Ersetze NUR den aktuellen Platzhalter durch die Benutzereingabe
        processedText = Replace(processedText, fullPlaceholder, userInstruction, 1, 1, vbTextCompare)
        
        ' Suche nach dem NÄCHSTEN Vorkommen für den nächsten Schleifendurchlauf
        startPos = InStr(1, processedText, "{{ASKINSTRUCTION", vbTextCompare)
    Loop


    ' 3. {{CLIPBOARD}}-Platzhalter behandeln
    If InStr(1, processedText, "{{CLIPBOARD}}", vbTextCompare) > 0 Then
        ' ... (Rest der Funktion bleibt unverändert)
        Dim clipboardContent As String
        clipboardContent = GetClipboardTextWithPreprocessing(removeSonderzeichen)
        processedText = Replace(processedText, "{{CLIPBOARD}}", clipboardContent, 1, -1, vbTextCompare)
    End If

    ' 4. {{TEXTSELECTED}}-Platzhalter behandeln
    If InStr(1, processedText, "{{TEXTSELECTED}}", vbTextCompare) > 0 Then
        ' ... (Rest der Funktion bleibt unverändert)
        Dim selectedText As String
        selectedText = Application.Selection.text
        If removeSonderzeichen Then
            selectedText = SonderzeichenEntfernen(selectedText)
        End If
        processedText = Replace(processedText, "{{TEXTSELECTED}}", selectedText, 1, -1, vbTextCompare)
    End If

    ' 5. {{TEXTBEFORE}}-Platzhalter behandeln (NEU: mit optionaler Längenangabe)
    Dim textBeforePlaceholderStart As Long, textBeforePlaceholderEnd As Long
    Dim fullTextBeforePlaceholder As String, textBeforeLengthStr As String, tempStrBefore As String
    Dim textBeforeLength As Long

    textBeforePlaceholderStart = InStr(1, processedText, "{{TEXTBEFORE", vbTextCompare)
    Do While textBeforePlaceholderStart > 0
        textBeforePlaceholderEnd = InStr(textBeforePlaceholderStart, processedText, "}}", vbTextCompare)
        If textBeforePlaceholderEnd = 0 Then Exit Do ' Unvollständigen Platzhalter ignorieren

        fullTextBeforePlaceholder = Mid(processedText, textBeforePlaceholderStart, textBeforePlaceholderEnd - textBeforePlaceholderStart + 2)
        
        ' Standardlänge festlegen
        textBeforeLength = 2000

        ' Prüfen, ob eine benutzerdefinierte Länge angegeben ist (z.B. :500)
        Dim colonPos As Long
        colonPos = InStr(1, fullTextBeforePlaceholder, ":", vbTextCompare)

        If colonPos > 0 Then
            ' Extrahiere den Teil nach dem Doppelpunkt und vor den "}}"
            tempStrBefore = Mid(fullTextBeforePlaceholder, colonPos + 1)
            textBeforeLengthStr = Left(tempStrBefore, Len(tempStrBefore) - 2)
            
            If IsNumeric(textBeforeLengthStr) Then
                textBeforeLength = CLng(Trim(textBeforeLengthStr)) ' Trim() entfernt Leerzeichen
            End If
        End If

        Dim textBefore As String
        Dim rngBefore As Range
        Dim startPosBefore As Long
        
        startPosBefore = Application.Selection.Start - textBeforeLength
        If startPosBefore < 0 Then startPosBefore = 0
        
        Set rngBefore = Application.ActiveDocument.Range(Start:=startPosBefore, End:=Application.Selection.Start)
        textBefore = rngBefore.text

        If removeSonderzeichen Then
            textBefore = SonderzeichenEntfernen(textBefore)
        End If

        ' Ersetze den spezifischen Platzhalter durch den abgerufenen Text
        processedText = Replace(processedText, fullTextBeforePlaceholder, textBefore, 1, 1, vbTextCompare)

        ' Suche nach dem nächsten Platzhalter für den nächsten Durchlauf
        textBeforePlaceholderStart = InStr(1, processedText, "{{TEXTBEFORE", vbTextCompare)
    Loop

    ' 6. {{TEXTAFTER}}-Platzhalter behandeln (NEU: mit optionaler Längenangabe)
    Dim textAfterPlaceholderStart As Long, textAfterPlaceholderEnd As Long
    Dim fullTextAfterPlaceholder As String, textAfterLengthStr As String, tempStrAfter As String
    Dim textAfterLength As Long

    textAfterPlaceholderStart = InStr(1, processedText, "{{TEXTAFTER", vbTextCompare)
    Do While textAfterPlaceholderStart > 0
        textAfterPlaceholderEnd = InStr(textAfterPlaceholderStart, processedText, "}}", vbTextCompare)
        If textAfterPlaceholderEnd = 0 Then Exit Do ' Unvollständigen Platzhalter ignorieren

        fullTextAfterPlaceholder = Mid(processedText, textAfterPlaceholderStart, textAfterPlaceholderEnd - textAfterPlaceholderStart + 2)
        
        ' Standardlänge festlegen
        textAfterLength = 2000
        
        ' Prüfen, ob eine benutzerdefinierte Länge angegeben ist (z.B. :500)
        Dim colonPosAfter As Long
        colonPosAfter = InStr(1, fullTextAfterPlaceholder, ":", vbTextCompare)

        If colonPosAfter > 0 Then
            ' Extrahiere den Teil nach dem Doppelpunkt und vor den "}}"
            tempStrAfter = Mid(fullTextAfterPlaceholder, colonPosAfter + 1)
            textAfterLengthStr = Left(tempStrAfter, Len(tempStrAfter) - 2)
            
            If IsNumeric(textAfterLengthStr) Then
                textAfterLength = CLng(Trim(textAfterLengthStr)) ' Trim() entfernt Leerzeichen
            End If
        End If

        Dim textAfter As String
        Dim rngAfter As Range
        Dim endPosAfter As Long
        
        endPosAfter = Application.Selection.End + textAfterLength
        If endPosAfter > Application.ActiveDocument.content.End Then
            endPosAfter = Application.ActiveDocument.content.End
        End If
        
        Set rngAfter = Application.ActiveDocument.Range(Start:=Application.Selection.End, End:=endPosAfter)
        textAfter = rngAfter.text

        If removeSonderzeichen Then
            textAfter = SonderzeichenEntfernen(textAfter)
        End If

        ' Ersetze den spezifischen Platzhalter durch den abgerufenen Text
        processedText = Replace(processedText, fullTextAfterPlaceholder, textAfter, 1, 1, vbTextCompare)

        ' Suche nach dem nächsten Platzhalter für den nächsten Durchlauf
        textAfterPlaceholderStart = InStr(1, processedText, "{{TEXTAFTER", vbTextCompare)
    Loop

    ' Gib den final bearbeiteten Text zurück.
    InjectPrompt = processedText
    
    Exit Function

ErrorHandler:
    MsgBox "Fehler " & Err.Number & ": " & Err.Description & " in InjectPrompt"
    InjectPrompt = promptText ' Im Fehlerfall den Originaltext zurückgeben
End Function

'------------------------------------------------------------------------------
' Funktion: ExtractModelName
' Zweck:    Sucht in einem Text nach dem {{MODEL:"..."}}-Platzhalter und
'           extrahiert den Modellnamen aus den Anführungszeichen.
' Parameter:
'   - promptText: Der zu durchsuchende rohe Prompt-Text.
' Gibt zurück: Den extrahierten Modellnamen (z.B. "gemma3:latest") oder
'              einen leeren String, wenn nichts gefunden wird.
'------------------------------------------------------------------------------
Public Function ExtractModelName(ByVal promptText As String) As String
    Dim startMarker As String
    Dim endMarker As String
    Dim startPos As Long
    Dim endPos As Long
    
    startMarker = "{{MODEL:"""  ' Sucht nach {{MODEL:"
    endMarker = """}}"          ' Sucht nach "}}
    
    ' Finde die Startposition des Markers
    startPos = InStr(1, promptText, startMarker, vbTextCompare)
    
    If startPos > 0 Then
        ' Finde die Endposition des Markers, beginnend nach dem Startmarker
        endPos = InStr(startPos + Len(startMarker), promptText, endMarker, vbTextCompare)
        
        If endPos > 0 Then
            ' Extrahiere den Text zwischen den Markern
            ExtractModelName = Mid(promptText, startPos + Len(startMarker), endPos - (startPos + Len(startMarker)))
            Exit Function ' Erfolgreich extrahiert, Funktion beenden
        End If
    End If
    
    ' Wenn nichts gefunden wurde, leeren String zurückgeben
    ExtractModelName = ""
End Function

'------------------------------------------------------------------------------
' Funktion: GetAllPromptCommands
' Zweck: Ruft alle verfügbaren Prompt-Befehle von der OpenWebUI API ab.
' Rückgabewert: Eine VBA.Collection mit allen Befehlsnamen (z.B. "/summarize").
'               Gibt bei einem Fehler eine leere Collection zurück.
'------------------------------------------------------------------------------
Public Function GetAllPromptCommands() As VBA.Collection
    Debug.Print "GetAllPromptComands startet"
    On Error GoTo ErrorHandler

    Dim request As Object
    Dim jsonResponse As String
    Dim commands As New VBA.Collection
    
    ' --- Cache-Buster, um sicherzustellen, dass die Daten frisch sind ---
    Dim cacheBuster As String
    cacheBuster = "cache=" & Round(Timer * 100, 0)
    
    Debug.Print OWUI_API_URL

    ' 1. HTTP-Anfrage erstellen
    Set request = CreateObject("MSXML2.XMLHTTP")
    request.Open "GET", OWUI_API_URL & "/api/v1/prompts/list" & "?" & cacheBuster, False
    request.setRequestHeader "Authorization", "Bearer " & OWUI_API_TOKEN
    'Debug.Print request
    request.send
    
    ' ===== NEUE DEBUG-ZEILEN START =====
    ' Gibt den HTTP-Status und die rohe Antwort des Servers im "Direktfenster" aus.
    'Debug.Print "HTTP Status: " & request.Status
    'Debug.Print "Server Response: " & request.responseText
    ' ===== NEUE DEBUG-ZEILEN ENDE =====

    ' 2. Antwort des Servers prüfen
    If request.Status = 200 Then
        jsonResponse = request.responseText
        
        ' 3. JSON manuell parsen, um alle "command"-Werte zu extrahieren
        '------------------------------------------------------------------
        Dim commandMarker As String
        Dim currentPos As Long
        Dim startPos As Long
        Dim endPos As Long
        
        commandMarker = """command"":"""
        currentPos = 1
        
        Do
            ' Finde den nächsten "command":"-Marker
            startPos = InStr(currentPos, jsonResponse, commandMarker, vbTextCompare)
            
            If startPos > 0 Then
                ' Position auf das erste Zeichen *nach* dem Marker setzen
                startPos = startPos + Len(commandMarker)
                
                ' Finde das schliessende Anführungszeichen
                endPos = InStr(startPos, jsonResponse, """", vbTextCompare)
                
                If endPos > 0 Then
                    ' Extrahiere den Befehl und füge ihn zur Collection hinzu
                    commands.Add Mid(jsonResponse, startPos, endPos - startPos)
                    ' Setze die Startposition für die nächste Suche
                    currentPos = endPos
                Else
                    Exit Do ' Kein schliessendes " gefunden, Schleife beenden
                End If
            Else
                Exit Do ' Kein "command":-Marker mehr gefunden
            End If
        Loop
        '------------------------------------------------------------------
    Else
        ' Im Fehlerfall wird eine leere Collection zurückgegeben
        MsgBox "Fehler beim Abrufen der Prompt-Liste." & vbCrLf & _
               "Status: " & request.Status & " - " & request.StatusText, vbExclamation
    End If

    Set GetAllPromptCommands = commands
    Set request = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Ein unerwarteter Fehler ist in GetAllPromptCommands aufgetreten: " & Err.Description, vbCritical
    Set GetAllPromptCommands = New VBA.Collection ' Leere Collection zurückgeben
    Set request = Nothing
End Function

'=== AI-Model Handling ===

'------------------------------------------------------------------------------
' Funktion: GetAllModels
' Zweck: Ruft die Liste der verfügbaren Modelle ab.
' Parameter:
'   - ApplyFilter (Optional, Boolean): Wenn True (Standard), wird die Liste
'     bereinigt (UUIDs entfernt, dedupliziert). Wenn False, werden alle
'     Modelle von der API zurückgegeben.
' Rückgabewert: Eine VBA.Collection mit allen Modell-IDs (z.B. "gemma3:27b").
'              Gibt bei einem Fehler eine leere Collection zurück.
'------------------------------------------------------------------------------
Public Function GetAllModels(Optional ByVal ApplyFilter As Boolean = True) As VBA.Collection
    On Error GoTo ErrorHandler

    Dim request As Object
    Dim jsonResponse As String
    Dim rawModels As New VBA.Collection      ' Temporäre Liste für alle Modelle von der API
    Dim finalModels As New VBA.Collection    ' Die finale, bereinigte Liste
    Dim seenNames As New VBA.Collection      ' Zum Prüfen auf exakte Duplikate

    ' 1. API-Anfrage senden und Rohdaten abrufen
    ' =================================================
    Set request = CreateObject("MSXML2.XMLHTTP")
    
    ' --- Cache-Buster hinzufügen, damit MSXML2.XMLHTTP nicht den Cache verwendet ---
    Dim cacheBuster As String
    cacheBuster = "cache=" & Round(Timer * 100, 0) ' Erzeugt eine quasi-zufällige Zahl
    
    request.Open "GET", OWUI_API_URL & "/ollama/api/tags" & "?" & cacheBuster, False
    request.setRequestHeader "Authorization", "Bearer " & OWUI_API_TOKEN
    request.send

    If request.Status = 200 Then
        jsonResponse = request.responseText
        ' Debug.Print jsonResponse

        ' 2. JSON parsen und ALLE Modelle in die "rawModels" Collection laden
        ' =================================================================
        Dim nameMarker As String, currentPos As Long, startPos As Long, endPos As Long
        nameMarker = """name"":"""
        currentPos = 1
        Do
            startPos = InStr(currentPos, jsonResponse, nameMarker, vbTextCompare)
            If startPos > 0 Then
                startPos = startPos + Len(nameMarker)
                endPos = InStr(startPos, jsonResponse, """", vbTextCompare)
                If endPos > 0 Then
                    rawModels.Add Mid(jsonResponse, startPos, endPos - startPos)
                    currentPos = endPos
                Else
                    Exit Do
                End If
            Else
                Exit Do
            End If
        Loop

        ' 3. Prüfen, ob der Filter angewendet werden soll
        ' =================================================================
        If ApplyFilter Then
            ' Filter- und angepasste Deduplizierungslogik anwenden
            Dim rawModelName As Variant
            
            For Each rawModelName In rawModels
                ' --- FILTER 1: UUID-ähnliche Namen überspringen ---
                If IsUUID(rawModelName) Then
                    ' Nichts tun, zum nächsten Modell springen
                Else
                    ' --- FILTER 2: Exakte Duplikate entfernen ---
                    
                    ' Prüfen, ob dieser exakte Name schon gesehen wurde.
                    On Error Resume Next ' Fehlerbehandlung für den Fall, dass der Key schon existiert
                    seenNames.Add rawModelName, rawModelName ' Füge den VOLLEN Namen als Key hinzu
                    
                    ' Wenn kein Fehler aufgetreten ist, ist der Name neu.
                    If Err.Number = 0 Then
                        finalModels.Add rawModelName ' Füge den Namen zur finalen Liste hinzu
                    End If
                    On Error GoTo 0 ' Fehlerbehandlung zurücksetzen
                End If
            Next rawModelName
            
            Set GetAllModels = finalModels ' Gib die bereinigte Liste zurück
        
        Else
            ' Filter ist ausgeschaltet: Gib die rohe, ungefilterte Liste zurück
            Set GetAllModels = rawModels
        End If

    Else
        MsgBox "Fehler beim Abrufen der Modell-Liste über /ollama/api/tags." & vbCrLf & _
               "Status: " & request.Status & " - " & request.StatusText, vbExclamation
        Set GetAllModels = New VBA.Collection ' Leere Collection bei Fehler zurückgeben
    End If

    Set request = Nothing
    Exit Function

ErrorHandler:
    MsgBox "Ein unerwarteter Fehler ist in GetAllModels aufgetreten: " & Err.Description, vbCritical
    Set GetAllModels = New VBA.Collection
    Set request = Nothing
End Function


'==============================================================================
' NEUE HILFSFUNKTION ZUR ERKENNUNG VON UUIDs
' wird von GetAllModels benötigt
'==============================================================================
Private Function IsUUID(ByVal text As String) As Boolean
    ' Einfache Heuristik zur Erkennung von UUID-ähnlichen Strings.
    ' Prüft, ob der String 36 Zeichen lang ist und Bindestriche enthält.
    If Len(text) = 36 And InStr(1, text, "-") > 0 Then
        IsUUID = True
    Else
        IsUUID = False
    End If
End Function

'=== Word Bridge Helper Functions ===

'------------------------------------------------------------------------------
' Function: GetClipboardTextWithPreprocessing
' Purpose: Gets text from clipboard with optional preprocessing
' Parameters:
'   - removeSonderzeichen: Whether to remove special characters
' Returns: Preprocessed text from clipboard
'------------------------------------------------------------------------------
Private Function GetClipboardTextWithPreprocessing(Optional removeSonderzeichen As Boolean = False) As String
    Dim zwischenablage As String
    zwischenablage = ClipBoard_GetData()

    Dim result As String
    result = FunctionZeilenumbruecheEntfernen(zwischenablage)
    result = RegelmaessigeOCRFehlerErsetzung(result)

    If removeSonderzeichen Then
        result = SonderzeichenEntfernen(result)
    End If

    GetClipboardTextWithPreprocessing = result
End Function


'------------------------------------------------------------------------------
' Funktion: ClipBoard_GetData
' Zweck:    Ruft den aktuellen Textinhalt aus der Windows-Zwischenablage ab.
'           Nutzt dafür direkte Windows-API-Aufrufe.
' Parameter:
'   - Keine
' Gibt zurück:
'   - String: Der Textinhalt der Zwischenablage. Wenn die Zwischenablage
'             leer ist, keinen Text enthält oder ein Fehler auftritt, wird
'             ein leerer String ("") zurückgegeben.
' Hinweise:
'   - Diese Funktion ist abhängig von Windows-API-Deklarationen
'     (OpenClipboard, GetClipboardData, etc.), die im Modul vorhanden
'     sein müssen.
'   - Bei einem Fehler (z.B. wenn die Zwischenablage von einer anderen
'     Anwendung blockiert wird) wird eine MsgBox angezeigt.
'------------------------------------------------------------------------------
Function ClipBoard_GetData() As String
    Dim hClipMemory As LongPtr 'LongPtr für 32/64-Bit-Kompatibilität
    Dim lpClipMemory As LongPtr 'LongPtr für 32/64-Bit-Kompatibilität
    Dim MyString As String
    Dim RetVal As Long
    
    ' Versuchen, die Zwischenablage zu öffnen. Wenn nicht erfolgreich, mit Meldung beenden.
    If OpenClipboard(0&) = 0 Then
        MsgBox "Zwischenablage konnte nicht aufgerufen werden, vielleicht ist sie durch ein anderes Programm geöffnet?"
        Exit Function
    End If
          
    ' Handle für die Daten im CF_TEXT Format aus der Zwischenablage abrufen.
    hClipMemory = GetClipboardData(CF_TEXT)
    
    ' Prüfen, ob ein gültiges Handle erhalten wurde.
    If hClipMemory = 0 Then
        ' Kein Text in der Zwischenablage oder ein anderer Fehler.
        ' Schliessen und mit leerem String beenden.
        GoTo OutOfHere
    End If

    ' Den Speicherblock sperren, um einen Zeiger auf die Daten zu erhalten.
    lpClipMemory = GlobalLock(hClipMemory)

    ' Prüfen, ob der Speicher erfolgreich gesperrt wurde.
    If lpClipMemory <> 0 Then
        ' Den Inhalt aus dem Speicherzeiger in einen VBA-String kopieren.
        MyString = Space$(lstrlen(lpClipMemory)) ' Speicherplatz genau zuweisen
        RetVal = lstrcpy(MyString, lpClipMemory)
        
        ' Speicher wieder freigeben.
        RetVal = GlobalUnlock(hClipMemory)
    Else
        MsgBox "Fehler: Der Speicher der Zwischenablage konnte nicht gesperrt werden."
    End If

OutOfHere:
    ' Zwischenablage immer schliessen, um sie für andere Programme freizugeben.
    RetVal = CloseClipboard()
    
    ' Den ermittelten String (kann auch leer sein) als Funktionsergebnis zurückgeben.
    ClipBoard_GetData = MyString
    
End Function


'------------------------------------------------------------------------------
' Funktion: FunctionZeilenumbruecheEntfernen
' Zweck:    Entfernt verschiedene Arten von Zeilenumbrüchen (Cr, Lf) aus einem
'           Text, um ihn in eine einzige, durchgehende Zeile umzuwandeln.
'           Die Funktion behandelt dabei intelligent Zeilenenden, die mit oder
'           ohne Bindestrich aufhören, um Wörter korrekt zu verbinden.
'           Zusätzlich werden Leerzeichen-Fehler und Null-Zeichen bereinigt.
' Parameter:
'   - textstelle: Der Eingabe-String, der Zeilenumbrüche enthalten kann.
' Gibt zurück:
'   - String:     Der bereinigte Text als einzelne Zeile ohne Zeilenumbrüche.
' Abhängigkeiten:
'   - Funktion "ersetze": Diese Hilfsfunktion für reguläre Ausdrücke muss
'     im Projekt verfügbar sein.
'------------------------------------------------------------------------------
Function FunctionZeilenumbruecheEntfernen(ByVal textstelle As String) As String
    
    'Allfällige Leerzeichen am ende der Zeile löschen
    textstelle = ersetze(textstelle, "\s+$", "")
    'Wenn Buchstabe, Ziffer oder Unterstrich am Ende der Zeile (ausser Bindestriche - davon gibt es mehrere Arten), Leerschlag hinzufügen
    textstelle = ersetze(textstelle, "([^-­])$", "$1 ")
    'Bindestrich am Ende der Zeile löschen
    textstelle = ersetze(textstelle, "-$", "")

    'Zeilenumbrüche (andere Art) entfernen
    textstelle = Replace(textstelle, Chr(10), "")
    'Carriage breaks (Zeilenumbrüche) entfernen
    textstelle = Replace(textstelle, Chr(13), "")
    
    'doppelte und dreifache Leerzeichen ersetzen
    textstelle = Replace(textstelle, "  ", " ")
    textstelle = Replace(textstelle, "   ", " ")
    
    'entferne Platzhalterquadrat für unbekanntes Zeichen
    textstelle = Replace(textstelle, ChrW(0), "")
    
    'Das Returnstatment funktioniert inden man den Namen der Funktion mit dem Return gleichsetzt
    FunctionZeilenumbruecheEntfernen = textstelle
    
End Function


'------------------------------------------------------------------------------
' Funktion: RegelmaessigeOCRFehlerErsetzung
' Zweck:    Bereinigt einen Text von einer vordefinierten Liste häufig
'           auftretender Fehler, die bei der automatischen Texterkennung (OCR)
'           entstehen. Der Fokus liegt auf typischen deutschen
'           Worterkennungsproblemen (z.B. "l" statt "I", "nrv" statt "rw").
' Parameter:
'   - textstelle: Der rohe Eingabe-String aus einer OCR-Quelle.
' Gibt zurück:
'   - String:     Der Text nach Anwendung der Korrekturregeln.
' Abhängigkeiten:
'   - Funktion "ersetze": Wird für einige Korrekturen benötigt.
' Hinweise:
'   - Die Funktion basiert auf einer statischen Liste von Ersetzungen, die
'     empirisch gesammelt wurden. Sie deckt nicht alle denkbaren
'     OCR-Fehler ab.
'------------------------------------------------------------------------------
Function RegelmaessigeOCRFehlerErsetzung(ByVal textstelle As String) As String

    'lch (Ich falsch geschrieben)
    textstelle = Replace(textstelle, "lch ", "Ich ")
    
    'fehlerhafte Interpretation des Anführungszeichens
    textstelle = ersetze(textstelle, ",,", Chr(34))
    
    'fehlerhafte Interpretation des Prozentzeichens
    textstelle = ersetze(textstelle, "o/o", "%")
    
    'regelmässige falsche Erkennung des w - test ohne Leerzeichen
    textstelle = Replace(textstelle, "nrv", "rw")
    textstelle = Replace(textstelle, "nru", "rw")
    textstelle = Replace(textstelle, "nry", "rw")
    
    'Ersetzung von OCR Fehlern: das grosse I wird fehlerhaft als l erkannt
    'In
    textstelle = Replace(textstelle, " ln ", " In ")
    textstelle = Replace(textstelle, " ln,", " In,")
    textstelle = Replace(textstelle, " ln.", " In.")
    'Information
    textstelle = Replace(textstelle, " lnformation ", " Information ")
    textstelle = Replace(textstelle, " lnformation,", " Information,")
    textstelle = Replace(textstelle, " lnformation.", " Information.")
    'Informationen
    textstelle = Replace(textstelle, " lnformationen ", " Informationen ")
    textstelle = Replace(textstelle, " lnformationen,", " Informationen,")
    textstelle = Replace(textstelle, " lnformationen.", " Informationen.")
    'Info
    textstelle = Replace(textstelle, " lnfo ", " Info ")
    textstelle = Replace(textstelle, " lnfo,", " Info,")
    textstelle = Replace(textstelle, " lnfo.", " Info.")
    'Infos
    textstelle = Replace(textstelle, " lnfos ", " Infos ")
    textstelle = Replace(textstelle, " lnfos,", " Infos,")
    textstelle = Replace(textstelle, " lnfos.", " Infos.")
    'Idee
    textstelle = Replace(textstelle, " ldee ", " Idee ")
    textstelle = Replace(textstelle, " ldee,", " Idee,")
    textstelle = Replace(textstelle, " ldee.", " Idee.")
    'Ideen
    textstelle = Replace(textstelle, " ldeen ", " Ideen ")
    textstelle = Replace(textstelle, " ldeen,", " Ideen,")
    textstelle = Replace(textstelle, " ldeen.", " Ideen.")
    'Insider
    textstelle = Replace(textstelle, " lnsider ", " Insider ")
    textstelle = Replace(textstelle, " lnsider,", " Insider,")
    textstelle = Replace(textstelle, " lnsider.", " Insider.")
    'Irrtum
    textstelle = Replace(textstelle, " lrrtum ", " Irrtum ")
    textstelle = Replace(textstelle, " lrrtum,", " Irrtum,")
    textstelle = Replace(textstelle, " lrrtum.", " Irrtum.")
    'Identifikation
    textstelle = Replace(textstelle, " ldentifikation ", " Identifikation ")
    textstelle = Replace(textstelle, " ldentifikation,", " Identifikation,")
    textstelle = Replace(textstelle, " ldentifikation.", " Identifikation.")
    'Instruktionen
    textstelle = Replace(textstelle, " lnstruktionen ", " Instruktionen ")
    textstelle = Replace(textstelle, " lnstruktionen,", " Instruktionen,")
    textstelle = Replace(textstelle, " lnstruktionen.", " Instruktionen.")
    'In (in am Satzanfang)
    textstelle = Replace(textstelle, " ln ", " In ")
    'Interesse
    textstelle = Replace(textstelle, " lnteresse ", " Interesse ")
    textstelle = Replace(textstelle, " lnteresse,", " Interesse,")
    textstelle = Replace(textstelle, " lnteresse.", " Interesse.")
    'Irgendwann
    textstelle = Replace(textstelle, " lrgendwann ", " Irgendwann ")
    textstelle = Replace(textstelle, " lrgendwann,", " Irgendwann,")
    textstelle = Replace(textstelle, " lrgendwann.", " Irgendwann.")
    'Inhalt
    textstelle = Replace(textstelle, " lnhalt ", " Inhalt ")
    textstelle = Replace(textstelle, " lnhalt,", " Inhalt,")
    textstelle = Replace(textstelle, " lnhalt.", " Inhalt.")
    'Inhaltlich
    textstelle = Replace(textstelle, " lnhaltlich ", " Inhaltlich ")
    textstelle = Replace(textstelle, " lnhaltlich,", " Inhaltlich,")
    textstelle = Replace(textstelle, " lnhaltlich.", " Inhaltlich.")
    'Interpretation
    textstelle = Replace(textstelle, " lnterpretation ", " Interpretation ")
    textstelle = Replace(textstelle, " lnterpretation,", " Interpretation,")
    textstelle = Replace(textstelle, " lnterpretation.", " Interpretation.")
    'Initiative
    textstelle = Replace(textstelle, " lnitiative ", " Initiative ")
    textstelle = Replace(textstelle, " lnitiative,", " Initiative,")
    textstelle = Replace(textstelle, " lnitiative.", " Initiative.")
    'Ihnen
    textstelle = Replace(textstelle, " lhnen ", " Ihnen ")
    textstelle = Replace(textstelle, " lhnen,", " Ihnen,")
    textstelle = Replace(textstelle, " lhnen.", " Ihnen.")
    'erwähnt
    textstelle = Replace(textstelle, " enruähnt  ", " erwähnt ")
    textstelle = Replace(textstelle, " enruähnt,", " erwähnt,")
    textstelle = Replace(textstelle, " enruähnt.", " erwähnt.")
    'erwähnen
    textstelle = Replace(textstelle, " enruähnen ", " erwähnen ")
    textstelle = Replace(textstelle, " enruähnen,", " erwähnen,")
    textstelle = Replace(textstelle, " enruähnen.", " erwähnen.")
    'erwähnen
    textstelle = Replace(textstelle, " enryähnen ", " erwähnen ")
    textstelle = Replace(textstelle, " enryähnen,", " erwähnen,")
    textstelle = Replace(textstelle, " enryähnen.", " erwähnen.")
    'erwähnen
    textstelle = Replace(textstelle, " enrvähnen ", " erwähnen ")
    textstelle = Replace(textstelle, " enrvähnen,", " erwähnen,")
    textstelle = Replace(textstelle, " enrvähnen.", " erwähnen.")
    'erwartet
    textstelle = Replace(textstelle, " enruartet ", " erwartet ")
    textstelle = Replace(textstelle, " enruartet,", " erwartet,")
    textstelle = Replace(textstelle, " enruartet.", " erwartet.")
    'normalerweise
    textstelle = Replace(textstelle, " normalenrueise ", " normalerweise ")
    textstelle = Replace(textstelle, " normalenrueise,", " normalerweise,")
    textstelle = Replace(textstelle, " normalenrueise.", " normalerweise.")
    'Inhaber
    textstelle = Replace(textstelle, " lnhaber ", " Inhaber ")
    textstelle = Replace(textstelle, " lnhaber,", " Inhaber,")
    textstelle = Replace(textstelle, " lnhaber.", " Inhaber.")
    'verwenden
    textstelle = Replace(textstelle, " venruenden ", " verwenden ")
    textstelle = Replace(textstelle, " venruenden ,", " verwenden,")
    textstelle = Replace(textstelle, " venruenden .", " verwenden.")
    'Immobilie
    textstelle = Replace(textstelle, " lmmobilie ", " Immobilie ")
    textstelle = Replace(textstelle, " lmmobilie,", " Immobilie,")
    textstelle = Replace(textstelle, " lmmobilie.", " Immobilie.")
    'Immobilien
    textstelle = Replace(textstelle, " lmmobilien ", " Immobilien ")
    textstelle = Replace(textstelle, " lmmobilien,", " Immobilien,")
    textstelle = Replace(textstelle, " lmmobilien.", " Immobilien.")
    'Investition
    textstelle = Replace(textstelle, " lnvestition ", " Investition ")
    textstelle = Replace(textstelle, " lnvestition,", " Investition,")
    textstelle = Replace(textstelle, " lnvestition.", " Investition.")
    'Investitionen
    textstelle = Replace(textstelle, " lnvestitionen ", " Investitionen ")
    textstelle = Replace(textstelle, " lnvestitionen,", " Investitionen,")
    textstelle = Replace(textstelle, " lnvestitionen.", " Investitionen.")
    'Irgendwo
    textstelle = Replace(textstelle, " lrgendwo ", " Irgendwo ")
    textstelle = Replace(textstelle, " lrgendwo,", " Irgendwo,")
    textstelle = Replace(textstelle, " lrgendwo.", " Irgendwo.")
    'lrgendwann
    textstelle = Replace(textstelle, " lrgendwann ", " Irgendwann ")
    textstelle = Replace(textstelle, " lrgendwann,", " Irgendwann,")
    textstelle = Replace(textstelle, " lrgendwann.", " Irgendwann.")
    'lrgendwie
    textstelle = Replace(textstelle, " lrgendwie ", " Irgendwie ")
    textstelle = Replace(textstelle, " lrgendwie,", " Irgendwie,")
    textstelle = Replace(textstelle, " lrgendwie.", " Irgendwie.")
    'Internet
    textstelle = Replace(textstelle, " lnternet ", " Internet ")
    textstelle = Replace(textstelle, " lnternet,", " Internet,")
    textstelle = Replace(textstelle, " lnternet.", " Internet.")
    'Internetseite
    textstelle = Replace(textstelle, " lnternetseite ", " Internetseite ")
    textstelle = Replace(textstelle, " lnternetseite,", " Internetseite,")
    textstelle = Replace(textstelle, " lnternetseite.", " Internetseite.")
    'Internetauftritt
    textstelle = Replace(textstelle, " lnternetauftritt ", " Internetauftritt ")
    textstelle = Replace(textstelle, " lnternetauftritt,", " Internetauftritt,")
    textstelle = Replace(textstelle, " lnternetauftritt.", " Internetauftritt.")
    'erwähnt
    textstelle = Replace(textstelle, " enrvähnt ", " erwähnt ")
    textstelle = Replace(textstelle, " enrvähnt,", " erwähnt,")
    textstelle = Replace(textstelle, " enrvähnt.", " erwähnt.")
    'Vorwurf
    textstelle = Replace(textstelle, " Vonryurf ", " Vorwurf ")
    textstelle = Replace(textstelle, " Vonryurf,", " Vorwurf,")
    textstelle = Replace(textstelle, " Vonryurf.", " Vorwurf.")
    'Vorwurf
    textstelle = Replace(textstelle, " Vonrvurf ", " Vorwurf ")
    textstelle = Replace(textstelle, " Vonrvurf,", " Vorwurf,")
    textstelle = Replace(textstelle, " Vonrvurf.", " Vorwurf.")
    'Ihrer
    textstelle = Replace(textstelle, " lhrer ", " Ihrer ")
    textstelle = Replace(textstelle, " lhrer,", " Ihrer,")
    textstelle = Replace(textstelle, " lhrer.", " Ihrer.")
    'Ihren
    textstelle = Replace(textstelle, " lhren ", " Ihren ")
    textstelle = Replace(textstelle, " lhren,", " Ihren,")
    textstelle = Replace(textstelle, " lhren.", " Ihren.")
    'normalerweise
    textstelle = Replace(textstelle, " normalenrveise ", " normalerweise ")
    textstelle = Replace(textstelle, " normalenrveise,", " normalerweise,")
    textstelle = Replace(textstelle, " normalenrveise.", " normalerweise.")
    'International
    textstelle = Replace(textstelle, " lnternational ", " International ")
    textstelle = Replace(textstelle, " lnternational,", " International,")
    textstelle = Replace(textstelle, " lnternational.", " International.")
    'Ingenieur
    textstelle = Replace(textstelle, " lngenieur ", " Ingenieur ")
    textstelle = Replace(textstelle, " lngenieur,", " Ingenieur,")
    textstelle = Replace(textstelle, " lngenieur.", " Ingenieur.")
    'Ingenieure
    textstelle = Replace(textstelle, " lngenieure ", " Ingenieure ")
    textstelle = Replace(textstelle, " lngenieure,", " Ingenieure,")
    textstelle = Replace(textstelle, " lngenieure.", " Ingenieure.")
    'Ingenieuren
    textstelle = Replace(textstelle, " lngenieuren ", " Ingenieuren ")
    textstelle = Replace(textstelle, " lngenieuren,", " Ingenieuren,")
    textstelle = Replace(textstelle, " lngenieuren.", " Ingenieuren.")
    'Investition
    textstelle = Replace(textstelle, " lnvestition ", " Investition ")
    textstelle = Replace(textstelle, " lnvestition,", " Investition,")
    textstelle = Replace(textstelle, " lnvestition.", " Investition.")
    'Investitionen
    textstelle = Replace(textstelle, " lnvestitionen ", " Investitionen ")
    textstelle = Replace(textstelle, " lnvestitionen,", " Investitionen,")
    textstelle = Replace(textstelle, " lnvestitionen.", " Investitionen.")
    'Indien
    textstelle = Replace(textstelle, " lndien ", " Indien ")
    textstelle = Replace(textstelle, " lndien,", " Indien,")
    textstelle = Replace(textstelle, " lndien.", " Indien.")
    'ltalien
    textstelle = Replace(textstelle, " ltalien ", " Italien ")
    textstelle = Replace(textstelle, " ltalien,", " Italien,")
    textstelle = Replace(textstelle, " ltalien.", " Italien.")
    'erwarteten
    textstelle = Replace(textstelle, " enrvarteten ", " erwarteten ")
    textstelle = Replace(textstelle, " enrvarteten,", " erwarteten,")
    textstelle = Replace(textstelle, " enrvarteten.", " erwarteten.")
    'erwartet
    textstelle = Replace(textstelle, " enrvartet ", " erwartet ")
    textstelle = Replace(textstelle, " enrvartet,", " erwartet,")
    textstelle = Replace(textstelle, " enrvartet.", " erwartet.")
    'Investor
    textstelle = Replace(textstelle, " lnvestor ", " Investor ")
    textstelle = Replace(textstelle, " lnvestor,", " Investor,")
    textstelle = Replace(textstelle, " lnvestor.", " Investor.")
    'Investoren
    textstelle = Replace(textstelle, " lnvestoren ", " Investoren ")
    textstelle = Replace(textstelle, " lnvestoren,", " Investoren,")
    textstelle = Replace(textstelle, " lnvestoren.", " Investoren.")
    'Inhaberaktien
    textstelle = Replace(textstelle, " lnhaberaktien ", " Inhaberaktien ")
    textstelle = Replace(textstelle, " lnhaberaktien,", " Inhaberaktien,")
    textstelle = Replace(textstelle, " lnhaberaktien.", " Inhaberaktien.")
    'Igor
    textstelle = Replace(textstelle, " lgor ", " Igor ")
    textstelle = Replace(textstelle, " lgor,", " Igor,")
    textstelle = Replace(textstelle, " lgor.", " Igor.")
    'Interessent
    textstelle = Replace(textstelle, " lnteressent ", " Interessent ")
    textstelle = Replace(textstelle, " lnteressent,", " Interessent,")
    textstelle = Replace(textstelle, " lnteressent.", " Interessent.")
    'Interessenten
    textstelle = Replace(textstelle, " lnteressenten ", " Interessenten ")
    textstelle = Replace(textstelle, " lnteressenten,", " Interessenten,")
    textstelle = Replace(textstelle, " lnteressenten.", " Interessenten.")
    'Inhalte
    textstelle = Replace(textstelle, " lnhalte ", " Inhalte ")
    textstelle = Replace(textstelle, " lnhalte,", " Inhalte,")
    textstelle = Replace(textstelle, " lnhalte.", " Inhalte.")
    'Industrie
    textstelle = Replace(textstelle, " lndustrie ", " Industrie ")
    textstelle = Replace(textstelle, " lndustrie,", " Industrie,")
    textstelle = Replace(textstelle, " lndustrie.", " Industrie.")
    'Investment
    textstelle = Replace(textstelle, " lnvestment ", " Investment ")
    textstelle = Replace(textstelle, " lnvestment,", " Investment,")
    textstelle = Replace(textstelle, " lnvestment.", " Investment.")
    'Investments
    textstelle = Replace(textstelle, " lnvestments ", " Investments ")
    textstelle = Replace(textstelle, " lnvestments,", " Investments,")
    textstelle = Replace(textstelle, " lnvestments.", " Investments.")
    'Ihre
    textstelle = Replace(textstelle, " lhre ", " Ihre ")
    textstelle = Replace(textstelle, " lhre,", " Ihre,")
    textstelle = Replace(textstelle, " lhre.", " Ihre.")
    'Ihr
    textstelle = Replace(textstelle, " lhr ", " Ihr ")
    textstelle = Replace(textstelle, " lhr,", " Ihr,")
    textstelle = Replace(textstelle, " lhr.", " Ihr.")
    'Ihrem
    textstelle = Replace(textstelle, " lhrem ", " Ihrem ")
    textstelle = Replace(textstelle, " lhrem,", " Ihrem,")
    textstelle = Replace(textstelle, " lhrem.", " Ihrem.")
    'Ihres
    textstelle = Replace(textstelle, " lhres ", " Ihres ")
    textstelle = Replace(textstelle, " lhres,", " Ihres,")
    textstelle = Replace(textstelle, " lhres.", " Ihres.")
    'Anrufe
    textstelle = Replace(textstelle, " Arwfe ", " Anrufe ")
    textstelle = Replace(textstelle, " Arwfe,", " Anrufe,")
    textstelle = Replace(textstelle, " Arwfe.", " Anrufe.")
    'Inventar
    textstelle = Replace(textstelle, " lnventar ", " Inventar ")
    textstelle = Replace(textstelle, " lnventar,", " Inventar,")
    textstelle = Replace(textstelle, " lnventar.", " Inventar.")
    'In
    textstelle = Replace(textstelle, " ln ", " In ")
    'Michèle
    textstelle = Replace(textstelle, " Michöle ", " Michèle ")
    textstelle = Replace(textstelle, " Michöle,", " Michèle,")
    textstelle = Replace(textstelle, " Michöle.", " Michèle.")
    'André
    textstelle = Replace(textstelle, " Andr6 ", " André ")
    textstelle = Replace(textstelle, " Andr6,", " André,")
    textstelle = Replace(textstelle, " Andr6.", " André.")
    'anrufen
    textstelle = Replace(textstelle, " arwfen ", " anrufen ")
    textstelle = Replace(textstelle, " arwfen,", " anrufen,")
    textstelle = Replace(textstelle, " arwfen.", " anrufen.")
    
    RegelmaessigeOCRFehlerErsetzung = textstelle

End Function

'------------------------------------------------------------------------------
' Private Funktion: ersetze
' Zweck:    Eine Hilfsfunktion, die eine "Suchen & Ersetzen"-Operation mithilfe
'           von regulären Ausdrücken (Regex) durchführt. Sie ist speziell für
'           den globalen (alle Vorkommen) und mehrzeiligen Modus konfiguriert,
'           was sie ideal für die Textbereinigung macht.
' Parameter:
'   - MyString: Der String, in dem Text ersetzt werden soll.
'   - Muster:   Das reguläre Ausdrucksmuster, nach dem gesucht wird.
'   - Ersatz:   Der Text, der das gefundene Muster ersetzen soll.
' Gibt zurück:
'   - String:   Der modifizierte String nach der Ersetzung.
' Hinweise:
'   - Nutzt "Late Binding" (CreateObject), daher ist kein fester Verweis auf
'     die "Microsoft VBScript Regular Expressions"-Bibliothek nötig.
'   - Die Eigenschaft "MultiLine = True" bewirkt, dass die Zeichen `^` und `$`
'     den Anfang und das Ende jeder einzelnen Zeile betreffen.
'------------------------------------------------------------------------------
Private Function ersetze(MyString As String, Muster As String, Ersatz As String) As String
    'Regex Funktion die bei der Funktion FunctionZeilenumbruecheEntfernen benötigt wird

    'Durch die folgende Initialisierung des regexObject bedarf es keiner Einbindung
    'über "Extras -> Verweise -> Microsoft VBScript Regular Expressions 5.5".
    Dim RegexObject As Object
    Set RegexObject = CreateObject("VBScript.RegExp")
    
    RegexObject.Global = True      ' Stellt sicher, dass ALLE Vorkommen ersetzt werden
    RegexObject.MultiLine = True   ' Erlaubt die Erkennung von Zeilenanfängen (^) und -enden ($)
    RegexObject.pattern = Muster   ' Das Suchmuster festlegen
    
    ' Die Ersetzung durchführen und das Ergebnis zurückgeben
    ersetze = RegexObject.Replace(MyString, Ersatz)

End Function

'------------------------------------------------------------------------------
' Function: SonderzeichenEntfernen
' Purpose: Removes special characters that could break JSON
' Parameters:
'   - inputString: The string to process
' Returns: String with special characters removed
'------------------------------------------------------------------------------
Private Function SonderzeichenEntfernen(ByVal inputString As String) As String
    ' Replace specific special characters with nothing
    inputString = Replace(inputString, """", "")
    inputString = Replace(inputString, "'", "")
    inputString = Replace(inputString, "/", "")

    SonderzeichenEntfernen = inputString
End Function


'------------------------------------------------------------------------------
' Prozedur: InsertTextIntoDocument
' Zweck:    Fügt den übergebenen Text an der aktuellen Cursor-Position ein.
'           Optional wird der Text vor dem Einfügen durch die Funktion
'           "BelegstelleHinzufuegen" modifiziert. In diesem Fall wird
'           anschliessend der Cursor neu positioniert.
' Parameter:
'   - text:            Der Text, der eingefügt werden soll.
'   - mitBelegstelle:  (Optional) Boolean. Wenn True, wird vor dem letzten
'                      Punkt ein Belegstellenverweis hinzugefügt.
'                      Standardwert ist False.
' Abhängigkeiten:
'   - Funktion "BelegstelleHinzufuegen"
'   - Prozedur "MoveCursorBackThreeSteps"
'------------------------------------------------------------------------------
Private Sub InsertTextIntoDocument(ByVal text As String, Optional ByVal mitBelegstelle As Boolean = False)
    ' Prüfen, ob der optionale Parameter auf True gesetzt ist.
    If mitBelegstelle Then
        ' FALL 1: Belegstelle wird hinzugefügt

        ' 1. ZUERST die Prozedur für die Belegstelle aufrufen.
        '    Jetzt gibt es keine Namenskollision mehr und VBA ruft die richtige FUNKTION auf.
        text = BelegstelleHinzufuegen(text)

        ' 2. DANACH den eigentlichen Text einfügen.
        Selection.TypeText text:=text

        ' 3. ZULETZT den Cursor bewegen.
        Call MoveCursorBackThreeSteps

    Else
        ' FALL 2: Es wird keine Belegstelle hinzugefügt

        ' In diesem Fall wird einfach nur der Text eingefügt.
        Selection.TypeText text:=text
    End If
End Sub


'------------------------------------------------------------------------------
' Funktion: BelegstelleHinzufuegen
' Zweck:    Findet den letzten Punkt in einem Text und fügt direkt davor
'           eine Klammer mit einem Belegstellenverweis ein. Das typische
'           Ergebnis ist die Umwandlung von "... Text." in
'           "... Text (Ziff. )."
' Parameter:
'   - textstelle:  Der Eingabe-String, der modifiziert werden soll.
'   - zitierweise: (Optional) Der Text für den Belegstellenverweis.
'                  Wird nichts angegeben, ist der Standardwert "Ziff.".
' Gibt zurück: Der modifizierte String mit dem eingefügten Belegstellenverweis.
'------------------------------------------------------------------------------
Function BelegstelleHinzufuegen(ByVal textstelle As String, Optional ByVal zitierweise As String = "Ziff.") As String
    '
    ' Fügt vor dem letzten Punkt eine Klammer mit der angegebenen oder der Standard-Zitierweise ein.
    '

    ' Durch die folgende Initialisierung des regexObject bedarf es keiner Einbindung
    ' über "Extras -> Verweise -> Microsoft VBScript Regular Expressions 5.5".
    Dim RegexObject As Object
    Set RegexObject = CreateObject("VBScript.RegExp")
    
    RegexObject.Global = False      ' Nicht nach mehreren Vorkommen suchen
    RegexObject.MultiLine = False     ' Nicht das Ende jeder Zeile als String-Ende behandeln
    RegexObject.pattern = "\.\W*$"    ' Entspricht dem letzten Punkt und optionalen Leer-/Umbruchzeichen am Ende

    ' Ersetze das gefundene Muster. Chr(160) ist ein geschütztes Leerzeichen,
    ' das einen Zeilenumbruch zwischen Zitierweise und Punkt verhindert.
    BelegstelleHinzufuegen = RegexObject.Replace(textstelle, " (" & zitierweise & Chr(160) & ").")

End Function


'------------------------------------------------------------------------------
' Prozedur: MoveCursorBackThreeSteps
' Zweck:    Positioniert den Cursor im aktiven Word-Dokument genau drei
'           Zeichen nach links von der aktuellen Position. Falls Text
'           markiert ist, wird die Markierung aufgehoben und der Cursor
'           an den Anfang der ursprünglichen Markierung gesetzt, bevor die
'           eigentliche Bewegung stattfindet.
' Parameter:
'   - Keine
' Gibt zurück:
'   - Keinen
'------------------------------------------------------------------------------
Function MoveCursorBackThreeSteps()
    ' This program moves the cursor three steps back in Word
    
    ' Deklariert eine Variable für die aktuelle Markierung/Auswahl
    Dim sel As Selection
    ' Weist der Variable die aktuelle Auswahl zu
    Set sel = Application.Selection
    
    ' Prüft, ob es sich nicht nur um einen Einfügepunkt handelt (d.h. ob Text ausgewählt ist)
    If sel.Type <> wdSelectionIP Then
        ' Hebt die Markierung auf und setzt den Cursor an deren Anfang
        sel.Collapse Direction:=wdCollapseStart
    End If
    
    ' Bewegt den Cursor drei Zeichen nach links
    sel.MoveLeft Unit:=wdCharacter, Count:=3
    
End Function

'Weitere Ideen:
'Tag {{QUOTE:"Belegstelle"}} um automatisch Belegstelle anzufügen und den Cursor an den richtigen Ort zu setzen,
'dass gerade die Belegstelle hinzugefügt werden kann.
