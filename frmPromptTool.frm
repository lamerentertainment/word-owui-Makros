' --- Belegstellen-Platzhalter, bei der die automatische Cursor-Positionierung erfolgt ---
Private Const CURSOR_TARGET_STRING As String = "(Ziff. )"

' --- Liste der Prompts für Kursiv-Formatierung als eine einzige Konstante ---
Private Const ITALIC_PROMPT_LIST As String = "/frage-und-aussagekonvertierung-mannlich,/frage-und-aussagekonvertierung-weiblich" ' Fügen Sie hier weitere hinzu

' --- Globale Variable für das Modul, um die Prompts zu speichern ---
Private promptsList As VBA.Collection

' --- Wird benötigt zum Speichern des Bereichs der letzten Antwort ---
Private lastResponseRange As Range


Private Sub CommandButton1_Click()

End Sub

Private Sub lblStatus_Click()

End Sub

'------------------------------------------------------------------------------
' Wird ausgeführt, wenn das Formular initialisiert wird.
'------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    Me.Caption = "Prompt Tool (Lade Konfiguration...)"
    lblStatus.Caption = "Lade verfügbare Modelle und Prompts..."
    DoEvents ' UI aktualisieren

    ' --- SCHRITT 1: Verfügbare Modelle von der API laden ---
    ' ========================================================
    Dim modelsList As VBA.Collection
    ' Debug.Print "Starte GetAllModels Funktion von Modul owui, um aktuelle Modellliste von OWUI zu erhalten"
    Set modelsList = owui.GetAllModels(ApplyFilter:=False) ' Alle Modelle von OWUI laden, gefiltert oder nicht
    ' Debug.Print "Modell Liste von OWUI erhalten"
    
    Me.cboModel.Clear ' Alte Einträge löschen
    
    If modelsList.Count > 0 Then
        ' ComboBox mit den abgerufenen Modellen füllen
        Dim modelName As Variant
        For Each modelName In modelsList
            Me.cboModel.AddItem modelName
        Next modelName
        
        ' Versuche, das Standardmodell aus owuiconfig auszuwählen
        On Error Resume Next ' Zur Sicherheit, falls das Standardmodell nicht existiert
        Me.cboModel.value = OWUI_MODEL
        If Err.Number <> 0 Then
            ' Wenn das Standardmodell nicht gefunden wurde, wähle das erste der Liste
            Me.cboModel.ListIndex = 0
        End If
        On Error GoTo 0
        
    Else
        ' Fallback, wenn keine Modelle geladen werden konnten
        Me.cboModel.AddItem "Keine Modelle gefunden"
        Me.cboModel.ListIndex = 0
        Me.cboModel.Enabled = False
    End If
    
    ' --- SCHRITT 2: Verfügbare Prompts von der API laden ---
    ' ========================================================
    Debug.Print "Prompts holen"
    Set promptsList = owui.GetAllPromptCommands()
    
    If promptsList.Count > 0 Then
        ' ComboBox mit den abgerufenen Prompts füllen
        Dim p As Variant
        For Each p In promptsList
            cboPrompts.AddItem p
        Next p
        cboPrompts.ListIndex = -1 ' Keine Vorauswahl
        lblStatus.Caption = "Wähle einen der " & promptsList.Count & " von OpenWebUI geladenen Prompts oder verfasse einen neuen."
    Else
        lblStatus.Caption = "Konnte keine Prompts laden."
        cboPrompts.AddItem "Keine Prompts gefunden."
        cboPrompts.Enabled = False
    End If
    
    Me.Caption = "Prompt Tool"
End Sub

'------------------------------------------------------------------------------
' Wird ausgeführt, wenn ein Prompt aus der Liste ausgewählt wird.
'------------------------------------------------------------------------------
Private Sub cboPrompts_Change()
    If cboPrompts.ListIndex = -1 Then Exit Sub ' Nichts tun, wenn die Auswahl gelöscht wird
    
    ' --- Benötigte Variablen deklarieren ---
    Dim selectedCommand As String
    Dim promptContent As String
    Dim modelToSelect As String  ' Variable aus der Logik hinzugefügt
    
    ' --- Bestehende Logik zum Laden des Prompt-Inhalts ---
    selectedCommand = cboPrompts.value
    lblStatus.Caption = "Lade Inhalt für '" & selectedCommand & "'..."
    DoEvents
    
    ' Verwende die bestehende Funktion, um den Inhalt des Prompts abzurufen
    promptContent = owui.GetPromptByCommandName(selectedCommand)
    
    ' --- Prüfen, ob der Inhalt erfolgreich geladen wurde ---
    If promptContent <> "" Then
        ' Den geladenen Prompt im Textfeld anzeigen
        txtPrompt.text = promptContent
        lblStatus.Caption = "Prompt geladen. Bereit zum Senden."
        
        ' 1. Den Modellnamen aus dem geladenen Prompt-Inhalt extrahieren
        modelToSelect = owui.ExtractModelName(promptContent)

        ' 2. Die Modell-ComboBox (cboModel) basierend auf dem Ergebnis aktualisieren
        If modelToSelect <> "" Then
            ' Versuche, das extrahierte Modell in der ComboBox zu selektieren
            On Error Resume Next ' Falls das Modell nicht in der Liste existiert
            Me.cboModel.value = modelToSelect
            
            ' Prüfen, ob ein Fehler aufgetreten ist (Modell nicht gefunden)
            If Err.Number <> 0 Then
                 ' Optional: Fehler behandeln und Standard auswählen
                MsgBox "Modell '" & modelToSelect & "' nicht in der Liste gefunden. Standard wird beibehalten.", vbExclamation
                Me.cboModel.ListIndex = 0 ' Wähle Standard (z.B. erstes Element)
            End If
            On Error GoTo 0 ' Fehlerbehandlung zurücksetzen
        Else
            ' 3. Fallback: Wenn der Prompt KEINEN {{MODEL}}-Tag hat,
            '    wähle das in owuiconfig.OWUI_MODEL definierte Standardmodell aus.
            On Error Resume Next ' Zur Sicherheit, falls das Standardmodell nicht in der Liste existiert
            Me.cboModel.value = OWUI_MODEL
            
            ' Wenn das Setzen des Standardmodells fehlschlägt (z.B. nicht in der Liste),
            ' dann als letzten Ausweg das erste Element der Liste wählen.
            If Err.Number <> 0 Then
                Me.cboModel.ListIndex = 0
            End If
            On Error GoTo 0 ' Fehlerbehandlung zurücksetzen
        End If
        
    Else
        ' Bestehende Logik für den Fehlerfall
        txtPrompt.text = "Fehler: Konnte Inhalt für '" & selectedCommand & "' nicht laden."
        lblStatus.Caption = "Fehler beim Laden des Prompts."
    End If
End Sub

'------------------------------------------------------------------------------
' Wird ausgeführt, wenn der "Senden"-Button geklickt wird.
'------------------------------------------------------------------------------
Private Sub btnSend_Click()
    Dim finalPrompt As String
    Dim model As String
    Dim result As String
    
    ' Startposition des Cursors VOR dem Streaming merken
    Dim startRange As Range
    Set startRange = Selection.Range
    
    model = cboModel.value
    
    If Trim(txtPrompt.text) = "" Then
        MsgBox "Bitte geben Sie einen Prompt ein oder wählen Sie einen aus der Liste.", vbExclamation
        Exit Sub
    End If

    ' Die Platzhalter erst jetzt, direkt vor dem Senden, ersetzen.
    finalPrompt = owui.InjectPrompt(txtPrompt.text)
    
    lblStatus.Caption = "Sende Anfrage an " & model & "..."
    Me.Repaint ' UI sofort aktualisieren
    
    ' Funktion aufrufen, um die Antwort direkt ins Word-Dokument zu streamen
    result = owui.StreamOWUIToWordWithModel(finalPrompt, model)
    
    ' Prüfen, ob eine gültige Antwort zurückkam (result enthält den VOLLEN Text)
    If result <> "" And result <> "Fehler bei der Anfrage" And result <> "Fehler" Then
        
        ' Den eingefügten Bereich NACH dem Streaming ermitteln und speichern
        ' Setzt das globale Range-Objekt auf den Bereich vom ursprünglichen Start
        ' bis zur neuen Endposition des Cursors.
        Set lastResponseRange = ActiveDocument.Range(startRange.Start, Selection.Range.End)
        
        lblStatus.Caption = "Antwort wurde eingefügt."
        
        ' === Backslashes in Anführungszeichen umwandeln (NUR IM NEUEN TEXT) ===
    
        Dim replaceRange As Range
        Set replaceRange = ActiveDocument.Range(lastResponseRange.Start, lastResponseRange.End)
        
        With replaceRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .text = "\"
            .Replacement.text = """"
            .Forward = True
            .Wrap = wdFindStop
            .Format = False
            .MatchCase = False
            .Execute Replace:=wdReplaceAll
        End With
        ' =====================================================================
        
        ' ======================================================================
        ' === NEU: Text vor dem letzten Doppelpunkt kursiv formatieren ===
        ' ======================================================================
        ' Prüfen, ob der richtige Prompt verwendet wurde
        If InStr(1, "," & ITALIC_PROMPT_LIST & ",", "," & Trim(Me.cboPrompts.value) & ",") > 0 Then
            
            Dim formatRange As Range
            Set formatRange = ActiveDocument.Range(lastResponseRange.Start, lastResponseRange.End)
            
            With formatRange.Find
                .ClearFormatting
                .text = ":"
                .Forward = False ' Rückwärtssuche
                .Wrap = wdFindStop
                
                ' Wenn .Execute erfolgreich ist, wird der formatRange auf den Fundort (den Doppelpunkt) reduziert.
                If .Execute Then
            
                    ' Erstelle den Bereich, der kursiv formatiert werden soll.
                    Dim italicRange As Range
                    Set italicRange = ActiveDocument.Range(lastResponseRange.Start, formatRange.Start)
                    italicRange.Font.Italic = True
                Else
                    ' MsgBox "FEHLER: Doppelpunkt wurde im Text nicht gefunden."
                End If
            End With
        End If
        ' ======================================================================
        
        ' === Cursor automatisch in Belegstelle positionieren ===
        Dim findRange As Range
        ' Die folgende Zeile stellt sicher, dass eine ECHTE, unabhängige Kopie
        ' des Bereichs erstellt wird, damit das Original "lastResponseRange" nicht verändert wird.
        Set findRange = ActiveDocument.Range(lastResponseRange.Start, lastResponseRange.End)
        
        ' Konfiguriere die Suche
        With findRange.Find
            .text = CURSOR_TARGET_STRING
            .Forward = True
            .Wrap = wdFindStop ' Nicht über das Ende des Bereichs hinaus suchen
            .MatchCase = True
            
            ' Führe die Suche aus
            If .Execute Then
                ' Wenn der Text gefunden wurde, ist findRange jetzt der gefundene Bereich.
                
                ' Ersetze den gefundenen Text durch die Version mit geschütztem Leerzeichen (Chr(160))
                findRange.text = "(Ziff." & Chr(160) & ")"
                
                ' Setze den Cursor an das Ende des gefundenen Bereichs...
                findRange.Select
                Selection.Collapse Direction:=wdCollapseEnd
                
                ' ...und bewege ihn ein Zeichen nach links (vor die schliessende Klammer).
                Selection.MoveLeft Unit:=wdCharacter, Count:=1
            End If
        End With
        ' ======================================================================
        
    Else
        lblStatus.Caption = "Fehler bei der Anfrage oder keine Antwort erhalten."
        ' Sicherstellen, dass kein ungültiger Range gespeichert wird
        Set lastResponseRange = Nothing
    End If

    lblStatus.Caption = "Antwort wurde eingefügt."
    
    ' --- Fokus zurück auf das Word-Dokument legen ---
    ' Fokus robust zurück auf das Word-Dokument legen
    Dim wordHWnd As LongPtr
    wordHWnd = FindWindowA("OpusApp", vbNullString)
    If wordHWnd <> 0 Then
        SetForegroundWindow wordHWnd
    End If
End Sub

'------------------------------------------------------------------------------
' Schliesst das Formular
'------------------------------------------------------------------------------
Private Sub btnClose_Click()
    Unload Me
End Sub

'------------------------------------------------------------------------------
' Hilfsprozedur, welche die eigentliche Logik zum Rückgängigmachen enthält.
' Diese ist KEINE Ereignisprozedur und kann daher Parameter haben.
' Der optionale Parameter showMessage definiert, ob die Statusmeldung "Die letzte
' Antwort wurde entfernt." erscheint, was bei Retry nicht erforderlich ist
'------------------------------------------------------------------------------
Private Sub PerformUndo(Optional ByVal showMessage As Boolean = True)
    ' Prüfen, ob ein gültiger Range gespeichert ist
    If lastResponseRange Is Nothing Then
        If showMessage Then
            MsgBox "Es gibt keine Aktion zum Rückgängigmachen.", vbInformation
        End If
        Exit Sub
    End If
    
    ' Prüfen, ob der Range im Dokument noch existiert
    On Error Resume Next
    Dim docStart As Long, docEnd As Long
    docStart = ActiveDocument.content.Start
    docEnd = ActiveDocument.content.End
    
    If lastResponseRange.Start >= docStart And lastResponseRange.End <= docEnd Then
        ' Der Range scheint gültig zu sein, also löschen
        lastResponseRange.Delete
        If showMessage Then
            lblStatus.Caption = "Die letzte Antwort wurde entfernt."
        End If
    Else
        If showMessage Then
            MsgBox "Der rückgängig zu machende Text wurde bereits aus dem Dokument entfernt.", vbExclamation
        End If
    End If
    On Error GoTo 0
    
    ' Den gespeicherten Range zurücksetzen
    Set lastResponseRange = Nothing
End Sub

'------------------------------------------------------------------------------
' Ereignisprozedur, die ausgeführt wird, wenn der Benutzer auf btnUndo klickt.
'------------------------------------------------------------------------------
Private Sub btnUndo_Click()
    ' Ruft die Hilfsprozedur auf und sagt ihr, ob sie eine gelöscht-Nachricht anzeigen soll.
    Call PerformUndo(showMessage:=True)
End Sub

'------------------------------------------------------------------------------
' Löscht die letzte Antwort und sendet den Prompt sofort erneut
'------------------------------------------------------------------------------
Private Sub btnRetry_Click()
    ' Prüfen, ob es überhaupt etwas zum Wiederholen gibt
    If lastResponseRange Is Nothing Then
        MsgBox "Es gibt keine vorherige Antwort, die wiederholt werden könnte.", vbInformation
        Exit Sub
    End If

    ' Schritt 1: Rufe die Hilfsprozedur auf, ABER ohne eine Statusmeldung anzuzeigen
    Call PerformUndo(showMessage:=False)

    ' Schritt 2: Führe die Logik zum Senden sofort erneut aus
    Call btnSend_Click
End Sub

