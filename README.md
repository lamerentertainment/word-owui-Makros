# OWUI-VBA-Word-Integration

## Zweck des Projekts

> Dieses Projekt stellt eine Reihe von VBA-Makros für Microsoft Word zur Verfügung, die eine direkte Integration mit einer OpenWebUI- und Ollama-Instanz ermöglichen. Benutzer können direkt aus Word heraus Anfragen an Sprachmodelle senden, vordefinierte Prompts nutzen, Text im Dokument verarbeiten und die Antworten des Modells direkt in ihr Dokument einfügen – wahlweise auch als Live-Stream.
>
> Das Hauptziel ist die Effizienzsteigerung bei der Texterstellung und -bearbeitung, indem leistungsstarke KI-Funktionen nahtlos in den täglichen Arbeitsablauf in Word integriert werden.

## Features

- **Direkte API-Kommunikation:** Sendet Anfragen von Word an eine Ollama-kompatible API.
- **Benutzerfreundliches UI:** Ein "Prompt Tool"-Formular zur einfachen Auswahl von Modellen und vordefinierten Prompts.
- **Dynamisches Prompt-System:** Prompts können dynamische Platzhalter verwenden, um kontextbezogene Informationen zu übergeben:
    - `{{CLIPBOARD}}`: Fügt Text aus der Zwischenablage ein.
    - `{{TEXTSELECTED}}`: Verwendet den aktuell markierten Text.
    - `{{TEXTBEFORE:n}}` / `{{TEXTAFTER:n}}`: Nutzt eine definierte Anzahl `n` von Zeichen vor oder nach dem Cursor.
    - `{{ASKINSTRUCTION:"Ihre Frage"}}`: Öffnet ein Dialogfenster, um zusätzliche Anweisungen vom Benutzer abzufragen.
    - `{{MODEL:"modellname"}}`: Definiert, welches Sprachmodell für diesen spezifischen Prompt verwendet werden soll.
- **Modell-Management:** Lädt automatisch alle verfügbaren Modelle von der Ollama-Instanz und macht sie im UI auswählbar.
- **Streaming-Support:** Antworten des Modells können Zeichen für Zeichen direkt in das Word-Dokument gestreamt werden.
- **Undo & Retry:** Eine soeben generierte Antwort kann per Knopfdruck entfernt (**Undo**) oder gelöscht und mit demselben Prompt neu generiert werden (**Retry**).
- **Intelligente Cursor-Positionierung:** Fügt der generierte Text eine definierte Belegstelle (z.B. `(Ziff. )`) ein, wird der Cursor automatisch an die passende Stelle zur direkten Eingabe der Ziffer platziert.
- **Nahtloser Arbeitsfluss:** Nach dem Einfügen einer Antwort wechselt der Fokus automatisch zurück in das Word-Dokument, sodass direkt weitergeschrieben werden kann.
- **Text-Vorverarbeitung:** Enthält Hilfsfunktionen zur Bereinigung von Text, z. B. zur Entfernung von Zeilenumbrüchen und zur Korrektur häufiger OCR-Fehler.
- **Sichere Konfiguration:** Trennt sensible Daten (API-Schlüssel, URL) in einer separaten Konfigurationsdatei.

## Projektstruktur

Das Projekt besteht aus den folgenden Hauptkomponenten:

- `owui.bas`: Das Herzstück des Projekts. Dieses Modul enthält die gesamte Logik für die API-Kommunikation, die Verarbeitung von Prompts, die Textmanipulation und die Steuerung des UIs.
- `owuiconfig.bas`: Eine separate Konfigurationsdatei, in der alle benutzerspezifischen Einstellungen (API-URL, Token, Standardmodell) gespeichert werden. Diese Datei sollte niemals in einem öffentlichen Repository veröffentlicht werden.
- `frmPromptTool.frm` / `frmPromptTool.frx`: Das UserForm, das als grafische Benutzeroberfläche für die Interaktion mit den Modellen und Prompts dient.

## Installation und Einrichtung

Folgen Sie diesen Schritten, um das Projekt in Ihrem Word zu installieren.

### 1. VBA-Module importieren

1.  Öffnen Sie Microsoft Word und drücken Sie `ALT` + `F11`, um den VBA-Editor zu öffnen.
2.  Klicken Sie im **Projekt-Explorer** (normalerweise links) mit der rechten Maustaste auf Ihr Projekt (z.B. *Normal* oder ein anderes Vorlagenprojekt).
3.  Wählen Sie **Datei importieren...**.
4.  Importieren Sie die folgenden Dateien:
    - `owui.bas`
    - `frmPromptTool.frm`

### 2. Konfigurationsmodul erstellen

1.  Klicken Sie im VBA-Editor erneut mit der rechten Maustaste auf Ihr Projekt und wählen Sie **Einfügen -> Modul**.
2.  Ein neues, leeres Modul wird erstellt. Benennen Sie dieses Modul im **Eigenschaftenfenster** (unten links) in `owuiconfig` um.
3.  Fügen Sie den folgenden Code in das `owuiconfig`-Modul ein:

    ```vba
    ' Modul: owuiconfig.bas
    ' ZWECK: Enthält alle benutzerspezifischen Konfigurationsvariablen und Einstellungen
    '        für die OWUI-Anwendung.
    ' WICHTIG: Befindet sich diese Datei in einem öffentlichen Repo, sollte sie zur .gitignore-Datei
    '          hinzugefügt werden, um das versehentliche Hochladen von privaten API-Schlüsseln
    '          zu verhindern.
    Option Explicit

    '=====================================================
    ' ANWENDUNGSKONFIGURATION (OWUI)
    '=====================================================

    ' API-Endpunkt und persönlicher Token
    ' Beispiel-URL für OpenWebUI: "http://deine-ip:8080"
    ' Beispiel-URL für Ollama: "http://localhost:11434"
    Public Const OWUI_API_URL As String = "http://IHRE_OPENWEBUI_URL"
    Public Const OWUI_API_TOKEN As String = "Ihr_OpenWebUI_API_Token"

    ' Das Modell, das standardmässig verwendet werden soll.
    ' Dieses Modell wird vorausgewählt, falls ein Prompt kein eigenes Modell spezifiziert.
    ' Beispiel: "llama3:latest"
    Public Const OWUI_MODEL As String = "gemma:latest"
    ```

### 3. Konfiguration anpassen

- **`OWUI_API_URL`**: Ersetzen Sie `http://IHRE_OPENWEBUI_URL` durch die vollständige URL Ihrer OpenWebUI- oder Ollama-Instanz.
- **`OWUI_API_TOKEN`**: Tragen Sie hier Ihren persönlichen API-Schlüssel für OpenWebUI ein.
- **`OWUI_MODEL`**: Geben Sie den Namen des Modells an, das standardmässig verwendet werden soll.

### 4. Wichtiger Sicherheitshinweis

Wenn Sie Ihr Word-Projekt oder Ihre `Normal.dotm` in einem Git-Repository verwalten, stellen Sie **unbedingt** sicher, dass die Konfigurationsdatei `owuiconfig.bas` von der Versionskontrolle ausgeschlossen wird. Fügen Sie den Namen der Datei zu Ihrer `.gitignore`-Datei hinzu, um zu verhindern, dass Ihr API-Token versehentlich veröffentlicht wird.

```gitignore
# .gitignore
owuiconfig.bas```

## Verwendung

### Starten des Prompt-Tools

Das Haupt-UI wird über das Makro `ShowPromptForm` gestartet. Sie können dieses Makro zur Symbolleiste für den Schnellzugriff hinzufügen, um es bequem aufrufen zu können.

1.  Führen Sie das Makro `ShowPromptForm` aus.
2.  Das "Prompt Tool"-Fenster öffnet sich. Es lädt automatisch alle verfügbaren Modelle und vordefinierten Prompts von Ihrer OpenWebUI-Instanz.

### Das Prompt-Tool-Fenster

- **Prompt-Auswahl (Dropdown):** Wählen Sie einen vordefinierten Prompt aus. Der Inhalt wird automatisch in das Textfeld geladen.
- **Modell-Auswahl (Dropdown):** Wählen Sie das gewünschte Sprachmodell aus. Wenn ein geladener Prompt einen `{{MODEL:"..."}}`-Tag enthält, wird das entsprechende Modell automatisch ausgewählt.
- **Prompt-Textfeld:** Hier können Sie einen eigenen Prompt verfassen oder einen geladenen Prompt anpassen.
- **Senden-Button:** Sendet den Prompt an das ausgewählte Modell. Die Antwort wird an der aktuellen Cursor-Position im Word-Dokument eingefügt.
- **Undo-Button:** Macht die letzte eingefügte Antwort rückgängig.
- **Retry-Button:** Entfernt die letzte Antwort und sendet den ursprünglichen Prompt sofort erneut an das Modell.
- **Schliessen-Button:** Schliesst das Fenster.

### Beispiel-Makros

Das Projekt enthält vordefinierte Makros wie `aiKonvertiereWeiblich` und `aiKonvertiereMaennlich`. Diese dienen als Beispiel dafür, wie man spezifische Aufgaben direkt an ein Makro binden kann, ohne das UI zu öffnen. Sie holen sich einen vordefinierten Prompt, füllen ihn mit Daten aus der Zwischenablage und fügen das formatierte Ergebnis mit einer Belegstelle in das Dokument ein.

## Zukünftige Ideen

- Implementierung eines `{{QUOTE:"Belegstelle"}}`-Tags, um das Hinzufügen von Belegstellen weiter zu automatisieren.
- Erweiterung der Fehlerbehandlung und des Benutzefeedbacks.

-----

## Autor

Jonas Achermann, Kriminalgericht Luzern (jonas.achermann@lu.ch)
