' Modul: owuiconfig.bas
' ZWECK: Enthält alle benutzerspezifischen Konfigurationsvariablen und Einstellungen
'          für die OWUI-Anwendung.
' WICHTIG: Befindet sich diese Datei in einem öffentlichen Repo, sollte sie zur .gitignore-Datei 
'          hinzugefügt werden, um das
'          versehentliche Hochladen von privaten API-Schlüsseln zu verhindern.
Option Explicit

'=====================================================
' ANWENDUNGSKONFIGURATION (OWUI)
'=====================================================

' API-Endpunkt und persönlicher Token
Private Const OWUI_API_URL As String = "URL HIER EINSETZEN"
Private Const OWUI_API_TOKEN As String = "API KEY HIER EINSETZEN"

' Das Modell, das für Anfragen verwendet wird.
Public Const OWUI_MODEL As String = "TAG DES BASISMODELL HIER EINSETZEN"
