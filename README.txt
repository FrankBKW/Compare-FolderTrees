===============================================================================

                        C O M P A R E - F O L D E R T R E E S

              Zwei Ordnerstrukturen vergleichen - HTML-Bericht und CSV

                                 Version 1.3

===============================================================================


  INHALT
  ------

    1 .... Was das Script macht
    2 .... Voraussetzungen
    3 .... Schnellstart
    4 .... Wie verglichen wird
    5 .... Die Ausgabe
    6 .... Alle Parameter
    7 .... Beispiele
    8 .... Was man wissen sollte
    9 .... Aenderungen


-------------------------------------------------------------------------------
  1   WAS DAS SCRIPT MACHT
-------------------------------------------------------------------------------

  Es vergleicht zwei Ordner samt allen Unterordnern und beantwortet drei Fragen:

      Was ist nur auf der einen Seite vorhanden?
      Was gibt es auf beiden Seiten, unterscheidet sich aber?
      Wo genau im Verzeichnisbaum stecken die Abweichungen?

  Heraus kommt ein HTML-Bericht zum Durchklicken und drei CSV-Dateien zum
  Weiterverarbeiten in Excel.

  Typische Einsaetze: Backup verifizieren, zwei Archivstaende abgleichen,
  eine Datenmigration kontrollieren, NAS gegen lokale Kopie pruefen.


-------------------------------------------------------------------------------
  2   VORAUSSETZUNGEN
-------------------------------------------------------------------------------

  Windows PowerShell 5.1 (auf jedem Windows 10 / 11 vorhanden) oder
  PowerShell 7 und neuer.

  Keine Zusatzmodule, keine Installation, keine Internetverbindung.
  Das Script ist eine einzelne Datei.


-------------------------------------------------------------------------------
  3   SCHNELLSTART
-------------------------------------------------------------------------------

  In der PowerShell:

      .\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup\Daten" -OpenReport

  Falls die Ausfuehrungsrichtlinie das Starten blockiert, ohne sie dauerhaft
  zu aendern:

      powershell -ExecutionPolicy Bypass -File .\Compare-FolderTrees.ps1 ^
                 -PfadA "D:\Daten" -PfadB "E:\Backup\Daten" -OpenReport

  Der Bericht landet standardmaessig unter:

      <Desktop>\Ordnervergleich\

  Mit -OpenReport oeffnet er sich am Ende automatisch im Browser.


-------------------------------------------------------------------------------
  4   WIE VERGLICHEN WIRD
-------------------------------------------------------------------------------

  Beide Baeume werden rekursiv eingelesen und ueber den RELATIVEN PFAD
  einander zugeordnet. Gross-/Kleinschreibung spielt bei der Zuordnung keine
  Rolle - genau wie im Windows-Dateisystem.

  Jedes Objekt bekommt einen von vier Status:

      Nur in A ........ existiert nur im ersten Ordner, fehlt in B
      Nur in B ........ existiert nur im zweiten Ordner, fehlt in A
      Unterschiedlich . beidseitig vorhanden, aber abweichend
      Identisch ....... Groesse und Aenderungsdatum stimmen ueberein

  Bei Dateien steht zusaetzlich dabei, WAS abweicht:

      Groesse ......... unterschiedliche Dateigroesse
      Aenderungsdatum . Zeitstempel weichen ueber die Toleranz hinaus ab
      Schreibweise .... gleicher Name, andere Gross-/Kleinschreibung
      Namenskodierung . gleicher Name, andere Unicode-Kodierung der Umlaute
      Inhalt (Hash) ... nur mit -CompareHash: gleiche Groesse, anderer Inhalt


  UMLAUTE: ZWEI SCHREIBWEISEN, DIESELBE DATEI
  ...........................................

  Ein "ue" kann im Dateisystem auf zwei Arten gespeichert sein:

      Mueller.pdf   77 252 108 ...      ein Zeichen (U+00FC)
                                        -> so schreibt Windows

      Mueller.pdf   77 117 776 108 ...  Grundzeichen + Trema (U+0075 U+0308)
                                        -> so schreiben macOS und viele NAS

  Beide sehen auf dem Bildschirm voellig identisch aus, sind als Zeichenkette
  aber verschieden. Ohne Gegenmassnahme taucht dieselbe Datei zweimal im
  Bericht auf - einmal als "Nur in A" und einmal als "Nur in B".

  Das Script normalisiert Namen deshalb vor dem Vergleich und weist eine
  abweichende Kodierung als Unterschied "Namenskodierung" aus.

  Relevant wird das, sobald eine der beiden Seiten von einem NAS, einem Mac
  oder aus einem Archiv stammt.


  VERZEICHNISSE WERDEN UEBER IHREN INHALT BEWERTET
  ................................................

  Der Status eines Ordners ergibt sich aus den enthaltenen Dateien, rekursiv
  ueber alle Ebenen - NICHT aus Gesamtgroesse und Dateianzahl. Der Grund:

      A\Ordner 1\doku.txt       (15 B)  |  B\Ordner 1\doku.txt       (15 B)
      A\Ordner 1\nur_a_sub.txt  ( 9 B)  |  B\Ordner 1\nur_b_sub.txt  ( 9 B)
      ---------------------------------  |  ---------------------------------
              2 Dateien, 24 Byte         |          2 Dateien, 24 Byte

  Beide Seiten kommen auf dieselbe Summe und dieselbe Anzahl. Der Inhalt ist
  trotzdem verschieden. Wuerde man nur Summe und Anzahl vergleichen, sagte der
  Bericht faelschlich "Identisch" - eine nur in A vorhandene Datei saehe dann
  so aus, als laege sie auch in B.

  Deshalb hat die Verzeichnistabelle drei Spalten, die alles darunter
  zusammenfassen:

      Fehlt in B ...... Dateien unterhalb dieses Ordners, die es in B nicht gibt
      Fehlt in A ...... Dateien unterhalb dieses Ordners, die es in A nicht gibt
      Abweichend ...... beidseitig vorhanden, aber unterschiedlich

  Damit arbeitet man sich von oben nach unten zur Fundstelle durch, ohne die
  komplette Dateiliste zu durchsuchen.


-------------------------------------------------------------------------------
  5   DIE AUSGABE
-------------------------------------------------------------------------------

  HTML-BERICHT
  ............

  Eine einzige Datei, komplett eigenstaendig - kein CDN, keine Nachladerei,
  keine Internetverbindung. Laesst sich per Mail verschicken oder auf einem
  abgeschotteten System oeffnen.

      Kennzahlen-Kacheln .... Gesamtzahlen je Status auf einen Blick
      Volltextsuche ......... ueber Pfad und Dateiname
      Filterbuttons ......... je Status, inklusive "Nur Abweichungen"
      Sortierbare Spalten ... Groessen numerisch, nicht alphabetisch
      Fixierte Kopfzeilen ... beim Scrollen sichtbar
      Druckbar .............. Bedienelemente werden im Druck ausgeblendet

  Fehlende Objekte sind doppelt gekennzeichnet: rotes "fehlt" in der
  betroffenen Spalte und ein farbiges Status-Badge mit farbigem Balken am
  Zeilenanfang.


  CSV-DATEIEN
  ...........

  Drei Dateien, UTF-8 mit BOM und Semikolon als Trennzeichen. Damit oeffnen
  sie in einem deutschsprachigen Excel direkt korrekt, inklusive Umlaute:

      Vergleich_Verzeichnisse_<Zeitstempel>.csv ..... nur Verzeichnisse
      Vergleich_Dateien_<Zeitstempel>.csv ........... nur Dateien
      Vergleich_Gesamt_<Zeitstempel>.csv ............ beides zusammen

  Groessen stehen doppelt drin: als Rohwert in Byte zum Rechnen und Sortieren,
  und formatiert (z. B. "12,34 MB") zum Lesen.


-------------------------------------------------------------------------------
  6   ALLE PARAMETER
-------------------------------------------------------------------------------

  -PfadA <Pfad>                 Erster Ordner (Referenz).           ERFORDERLICH
  -PfadB <Pfad>                 Zweiter Ordner (Vergleich).         ERFORDERLICH

  -OutputFolder <Pfad>          Zielordner fuer HTML und CSV.
                                Standard: <Desktop>\Ordnervergleich

  -Filter <Muster>              Nur bestimmte Dateien, z. B. '*.pdf'.
                                Standard: *  (alle)

  -ExcludeFolder <Namen>        Ordnernamen ueberspringen, Wildcards erlaubt,
                                mehrere durch Komma getrennt.
                                Wirkt nur auf Ordner, nicht auf Dateinamen.

  -TimeToleranceSeconds <Zahl>  Toleranz beim Datumsvergleich.
                                Standard: 2

  -CompareHash                  Zusaetzlich Pruefsumme vergleichen, aber nur
                                bei gleicher Groesse. Liest jede Datei
                                vollstaendig - bei viel Datenmenge langsam.

  -HashAlgorithm <Name>         MD5 (Standard), SHA1 oder SHA256.

  -CsvDelimiter <Zeichen>       Trennzeichen der CSV-Dateien.
                                Standard: ;   (fuer englisches Excel: ,)

  -OpenReport                   HTML-Bericht am Ende automatisch oeffnen.


-------------------------------------------------------------------------------
  7   BEISPIELE
-------------------------------------------------------------------------------

  Backup gegen Original pruefen und den Bericht direkt anzeigen:

      .\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup\Daten" -OpenReport


  Nur PDF-Dateien vergleichen, Quelle liegt auf dem NAS:

      .\Compare-FolderTrees.ps1 -PfadA "D:\Rechnungen" -PfadB "\\nas\archiv\rechnungen" -Filter '*.pdf'


  Projektordner abgleichen und Arbeitsverzeichnisse ausblenden, mit Pruefsumme:

      .\Compare-FolderTrees.ps1 "D:\Projekte" "\\nas\projekte" -ExcludeFolder '.git','node_modules','bin','obj' -CompareHash


  Ergebnis im eigenen Script weiterverwenden - das Script liefert ein Objekt
  mit den Ausgabepfaden zurueck:

      $r = .\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup" -OutputFolder "C:\Reports"
      Start-Process $r.HtmlReport


-------------------------------------------------------------------------------
  8   WAS MAN WISSEN SOLLTE
-------------------------------------------------------------------------------

  GROSS-/KLEINSCHREIBUNG
      Dateien werden case-insensitiv zugeordnet, weil Windows das ebenso
      handhabt. Weicht die Schreibweise ab (DATEI.TXT gegen datei.txt), gilt
      das als Unterschied "Schreibweise"; der Bericht zeigt beide Varianten.

  VERKNUEPFTE ORDNER (JUNCTIONS, SYMLINKS)
      PowerShell 5.1 folgt Verknuepfungen beim rekursiven Einlesen NICHT. Ihr
      Inhalt fehlt damit im Vergleich, obwohl der Explorer ihn anzeigt.
      Betroffene Ordner werden unter "Hinweise" im Bericht aufgefuehrt - dort
      nachsehen, wenn ein ganzer Teilbaum unerwartet fehlt.

  PFADLAENGE
      Windows PowerShell 5.1 kann Pfade ueber 260 Zeichen nicht lesen. Solche
      Faelle verschwinden nicht stillschweigend, sondern erscheinen als
      Lesefehler unten im HTML-Bericht. Unter PowerShell 7+ oder mit
      aktiviertem LongPathsEnabled tritt das Problem nicht auf.

  AENDERUNGSDATUM
      FAT und exFAT speichern Zeitstempel nur in 2-Sekunden-Schritten - daher
      die Standardtoleranz von 2 Sekunden. Bei Vergleichen ueber Zeitzonen
      hinweg oder nach einem Kopiervorgang ohne Zeitstempel-Uebernahme kann
      ein hoeherer Wert sinnvoll sein.

  GROSSE DATENBESTAENDE
      Alle Zeilen stecken direkt in der HTML-Datei. Ab etwa 100.000 Dateien
      wird der Bericht entsprechend gross und traege. Dann besser mit -Filter
      eingrenzen oder gleich mit den CSV-Dateien arbeiten.

  -COMPAREHASH
      Liest jede Datei vollstaendig ein. Ohne den Schalter wird nur ueber
      Groesse und Aenderungsdatum verglichen - das reicht in den meisten
      Faellen und ist um ein Vielfaches schneller.


-------------------------------------------------------------------------------
  9   AENDERUNGEN
-------------------------------------------------------------------------------

  VERSION 1.3

      NEU - Warnung bei unvollstaendigem Einlesen. Konnte eine Seite nicht
      komplett gelesen werden, steht das jetzt als roter Kasten oben im Bericht
      und in der Konsolenausgabe: welche Seite betroffen ist und dass deren
      Dateien faelschlich als "nur auf der anderen Seite" erscheinen koennen.

      Das ist die haeufigste Ursache fuer scheinbar falsche Ergebnisse. Pfade
      ueber 260 Zeichen treffen nur die tiefer verschachtelte Seite - dieselben
      Dateien sind unter dem kuerzeren Pfad der anderen Seite lesbar und stehen
      dann dort allein.

  VERSION 1.2.1

      KORREKTUR - Das Initialisierungsskript des HTML-Berichts lief auch ueber
      den Hinweis-Kasten, der weder Suchfeld noch Filter besitzt, und brach
      dort mit einem Fehler ab. Weil der Kasten als letztes Element steht,
      waren die Tabellen zu dem Zeitpunkt bereits verdrahtet - die Filter
      funktionierten also. Stuende der Kasten weiter oben, waeren saemtliche
      Filterbuttons wirkungslos gewesen.

  VERSION 1.2

      KORREKTUR - Dateinamen mit Umlauten oder Akzenten wurden je nach
      Unicode-Kodierung als zwei verschiedene Dateien behandelt. Dieselbe Datei
      erschien dadurch doppelt im Bericht: einmal als "Nur in A" und einmal als
      "Nur in B". Namen werden jetzt vor dem Vergleich normalisiert, eine
      abweichende Kodierung erscheint als Unterschied "Namenskodierung".

      NEU - Verknuepfte Ordner (Junctions, Symlinks) werden unter "Hinweise"
      im Bericht aufgefuehrt. PowerShell 5.1 folgt ihnen nicht, ihr Inhalt
      fehlte bisher stillschweigend im Vergleich.

  VERSION 1.1

      KORREKTUR - Verzeichnisse wurden nur ueber Gesamtgroesse und Anzahl
      verglichen. Zwei Ordner mit zufaellig gleicher Summe, aber
      unterschiedlichen Dateien, galten dadurch als "Identisch". Eine nur in A
      vorhandene Datei sah damit so aus, als laege sie auch in B. Der Status
      wird jetzt aus dem tatsaechlichen Inhalt abgeleitet.

      NEU - Spalten "Fehlt in B", "Fehlt in A" und "Abweichend" je Verzeichnis.

      NEU - Abweichende Gross-/Kleinschreibung wird als Unterschied
      ausgewiesen, mit beiden Schreibweisen im Bericht.

      KORREKTUR - -ExcludeFolder hat auch auf Dateinamen gefiltert.

      KORREKTUR - Laufwerkstammverzeichnisse wie D:\ wurden auf D: gekuerzt.
      Das meint unter Windows das AKTUELLE Verzeichnis dieses Laufwerks und
      konnte damit den falschen Ordner einlesen.

  VERSION 1.0

      Erste Fassung.


===============================================================================
