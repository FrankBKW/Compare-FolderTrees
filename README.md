# Compare-FolderTrees

PowerShell-Script zum Vergleich zweier Ordnerstrukturen inklusive aller Unterordner.
Ergebnis: ein interaktiver HTML-Bericht und drei CSV-Dateien.

Läuft mit Windows PowerShell 5.1 (Bordmittel, keine Module nötig) und PowerShell 7+.

## Schnellstart

```powershell
.\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup\Daten" -OpenReport
```

Ohne Anpassung der Ausführungsrichtlinie:

```powershell
powershell -ExecutionPolicy Bypass -File .\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup\Daten" -OpenReport
```

## Funktionsweise

Beide Bäume werden rekursiv eingelesen und über den **relativen Pfad** einander
zugeordnet (Groß-/Kleinschreibung wird ignoriert, wie im Windows-Dateisystem).
Jedes Objekt bekommt einen von vier Status:

| Status            | Bedeutung                                                |
|-------------------|----------------------------------------------------------|
| `Nur in A`        | existiert nur im ersten Ordner — fehlt in B              |
| `Nur in B`        | existiert nur im zweiten Ordner — fehlt in A             |
| `Unterschiedlich` | auf beiden Seiten vorhanden, aber abweichend             |
| `Identisch`       | Größe und Änderungsdatum (optional Hash) stimmen überein |

Bei Dateien wird zusätzlich ausgewiesen, *was* abweicht: Größe, Änderungsdatum,
Schreibweise, Namenskodierung oder — mit `-CompareHash` — der Inhalt.

### Umlaute: zwei Schreibweisen, dieselbe Datei

Ein „ü" kann im Dateisystem auf zwei Arten gespeichert sein:

```
Müller.pdf   →  77 252 108 ...    ü als ein Zeichen (U+00FC)          Windows
Müller.pdf   →  77 117 776 108 …  ü als u + Trema (U+0075 U+0308)     macOS, viele NAS
```

Beide sehen identisch aus, sind als Zeichenkette aber verschieden. Ohne
Gegenmaßnahme taucht dieselbe Datei zweimal im Bericht auf — einmal als
`Nur in A` und einmal als `Nur in B`. Das Script normalisiert Namen deshalb vor
dem Vergleich (Unicode Form C) und weist eine abweichende Kodierung als
Unterschied `Namenskodierung` aus.

Praktisch relevant wird das, sobald eine der beiden Seiten von einem NAS, einem
Mac oder aus einem Archiv stammt.

### Verzeichnisse werden über ihren Inhalt bewertet

Der Status eines Ordners ergibt sich aus den **enthaltenen Dateien**, rekursiv über
alle Ebenen — nicht aus Gesamtgröße und Dateianzahl. Das ist wichtig, weil zwei
Ordner zufällig gleich groß sein und trotzdem völlig verschiedene Dateien enthalten
können:

```
A\Ordner 1\doku.txt        (15 B)      B\Ordner 1\doku.txt        (15 B)
A\Ordner 1\nur_a_sub.txt   ( 9 B)      B\Ordner 1\nur_b_sub.txt   ( 9 B)
                    2 Dateien, 24 B                        2 Dateien, 24 B
```

Beide Seiten kommen auf 2 Dateien und 24 Byte — der Inhalt ist trotzdem
verschieden. Der Ordner wird korrekt als `Unterschiedlich` gemeldet, mit dem
Hinweis „1 fehlt in B, 1 fehlt in A".

Deshalb hat die Verzeichnistabelle drei zusätzliche Spalten, die jeweils alle
Ebenen darunter zusammenfassen:

| Spalte       | Bedeutung                                                   |
|--------------|-------------------------------------------------------------|
| `Fehlt in B` | Dateien unterhalb dieses Ordners, die es in B nicht gibt    |
| `Fehlt in A` | Dateien unterhalb dieses Ordners, die es in A nicht gibt    |
| `Abweichend` | Dateien, die es beidseitig gibt, die sich aber unterscheiden |

So findet man den Ort einer Abweichung von oben nach unten, ohne die komplette
Dateiliste durchsuchen zu müssen.

## Ausgabe

### 1. Verzeichnisse

Relativer Pfad, Größe A/B (Summe aller enthaltenen Dateien, rekursiv), Differenz,
Anzahl Dateien A/B sowie die drei Abweichungsspalten.

### 2. Dateien

Dateiname, Verzeichnis, Größe A/B, Änderungsdatum A/B, Art der Abweichung.

Fehlende Objekte sind doppelt gekennzeichnet: rotes **fehlt** in der betroffenen
Spalte plus farbiges Status-Badge mit farbigem Balken am Zeilenanfang.

### HTML-Bericht

Eine einzelne Datei ohne externe Abhängigkeiten (kein CDN, keine Internetverbindung
nötig) — damit auch per Mail versendbar oder auf abgeschotteten Systemen nutzbar:

* Kennzahlen-Kacheln mit den Gesamtzahlen je Status
* Volltextsuche über Pfad und Dateiname
* Filterbuttons je Status, inklusive **Nur Abweichungen**
* sortierbare Spalten — Größen und Anzahlen numerisch, nicht alphabetisch
* fixierte Tabellenköpfe, druckfreundliches Layout

### CSV-Dateien

Drei Dateien, UTF-8 mit BOM und `;` als Trennzeichen — öffnen damit direkt korrekt
in einem deutschsprachigen Excel:

| Datei                           | Inhalt                               |
|---------------------------------|--------------------------------------|
| `Vergleich_Verzeichnisse_*.csv` | nur Verzeichnisse                    |
| `Vergleich_Dateien_*.csv`       | nur Dateien                          |
| `Vergleich_Gesamt_*.csv`        | beides mit vereinheitlichten Spalten |

Größen stehen jeweils als Rohwert in Bytes **und** formatiert (`12,34 MB`) in der
Tabelle, damit sich in Excel rechnen und sortieren lässt, ohne die Lesbarkeit zu
verlieren.

## Parameter

| Parameter               | Standard                    | Zweck                                            |
|-------------------------|-----------------------------|--------------------------------------------------|
| `-PfadA`                | *erforderlich*              | erster Ordner (Referenz)                         |
| `-PfadB`                | *erforderlich*              | zweiter Ordner (Vergleich)                       |
| `-OutputFolder`         | `<Desktop>\Ordnervergleich` | Zielordner für HTML und CSV                      |
| `-Filter`               | `*`                         | Dateifilter, z. B. `*.pdf`                       |
| `-ExcludeFolder`        | *leer*                      | Ordnernamen überspringen, Wildcards erlaubt      |
| `-TimeToleranceSeconds` | `2`                         | Toleranz beim Datumsvergleich (FAT/NTFS-Rundung) |
| `-CompareHash`          | aus                         | zusätzlich Prüfsumme, nur bei gleicher Größe     |
| `-HashAlgorithm`        | `MD5`                       | `MD5`, `SHA1` oder `SHA256`                      |
| `-CsvDelimiter`         | `;`                         | auf `,` setzen für englischsprachiges Excel      |
| `-OpenReport`           | aus                         | HTML-Bericht am Ende öffnen                      |

## Beispiele

Nur PDFs vergleichen und den Bericht direkt öffnen:

```powershell
.\Compare-FolderTrees.ps1 -PfadA "D:\Rechnungen" -PfadB "\\nas\archiv\rechnungen" -Filter '*.pdf' -OpenReport
```

Quellcode-Ordner vergleichen, Arbeitsverzeichnisse ausblenden, Inhalte per Hash prüfen:

```powershell
.\Compare-FolderTrees.ps1 "D:\Projekte" "\\nas\projekte" -ExcludeFolder '.git','node_modules','bin','obj' -CompareHash
```

Backup verifizieren und im Script weiterverarbeiten — das Script gibt die
Ausgabepfade als Objekt zurück:

```powershell
$r = .\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup" -OutputFolder "C:\Reports"
if ($r.AnzahlDateien -gt 0) { Start-Process $r.HtmlReport }
```

## Hinweise

* **Groß-/Kleinschreibung:** Dateien werden case-insensitiv zugeordnet, weil
  Windows das ebenso handhabt. Weicht die Schreibweise ab (`DATEI.TXT` gegen
  `datei.txt`), gilt das als Unterschied `Schreibweise`; der Bericht zeigt dann
  beide Varianten nebeneinander.
* **Verknüpfte Ordner:** Junctions und Symlinks verfolgt PowerShell 5.1 beim
  rekursiven Einlesen nicht. Ihr Inhalt fehlt damit im Vergleich, obwohl der
  Explorer ihn anzeigt. Betroffene Ordner stehen unter „Hinweise" im Bericht —
  dort nachsehen, wenn ein ganzer Teilbaum unerwartet fehlt.
* **Pfadlänge:** Windows PowerShell 5.1 kann Pfade über 260 Zeichen nicht lesen.
  Solche Fälle werden nicht stillschweigend übergangen, sondern als Lesefehler
  unten im HTML-Bericht aufgeführt. Unter PowerShell 7+ oder mit aktiviertem
  `LongPathsEnabled` tritt das Problem nicht auf.
* **Änderungsdatum:** FAT/exFAT speichert Zeitstempel in 2-Sekunden-Schritten.
  Deshalb die Standardtoleranz von 2 Sekunden — bei Vergleichen über Zeitzonen
  oder nach einem Kopiervorgang mit Zeitstempel-Verlust ggf. höher setzen.
* **Große Bäume:** Alle Zeilen werden in die HTML-Datei eingebettet. Ab etwa
  100.000 Dateien wird der Bericht entsprechend groß und träge — dann besser mit
  `-Filter` eingrenzen oder die CSV-Dateien auswerten.
* **`-CompareHash`** liest jede Datei vollständig. Bei großen Datenmengen sollte
  man das gezielt einsetzen; ohne den Schalter wird nur über Größe und
  Änderungsdatum verglichen.
* **`-ExcludeFolder`** wirkt ausschließlich auf Ordnernamen, nicht auf Dateinamen.
  Zum Filtern von Dateien ist `-Filter` zuständig.

## Änderungen

### Version 1.2.1

* **Korrektur:** Das Initialisierungsskript des HTML-Berichts lief auch über den
  Hinweis-Kasten, der weder Suchfeld noch Filter besitzt, und brach dort mit
  einem Fehler ab. Weil der Kasten als letztes Element steht, waren die Tabellen
  zu dem Zeitpunkt bereits verdrahtet — die Filter funktionierten also. Stünde
  der Kasten weiter oben, wären sämtliche Filterbuttons wirkungslos gewesen.

### Version 1.2

* **Korrektur:** Dateinamen mit Umlauten oder Akzenten wurden je nach
  Unicode-Kodierung als zwei verschiedene Dateien behandelt. Dieselbe Datei
  erschien dadurch doppelt im Bericht — einmal als `Nur in A`, einmal als
  `Nur in B`. Namen werden jetzt vor dem Vergleich normalisiert; eine
  abweichende Kodierung erscheint als Unterschied `Namenskodierung`.
* Ordner-Verknüpfungen (Junctions, Symlinks) werden von PowerShell 5.1 nicht
  verfolgt. Ihr Inhalt fehlte bisher stillschweigend im Vergleich, obwohl er im
  Explorer sichtbar ist. Solche Ordner werden jetzt unter „Hinweise" im Bericht
  aufgeführt.

### Version 1.1

* **Korrektur:** Verzeichnisse wurden nur über Gesamtgröße und Dateianzahl
  verglichen. Zwei Ordner mit zufällig gleicher Summe, aber unterschiedlichen
  Dateien, wurden dadurch fälschlich als `Identisch` gemeldet — eine nur in A
  vorhandene Datei sah damit so aus, als läge sie auch in B. Der Status wird jetzt
  aus dem tatsächlichen Inhalt abgeleitet.
* Neue Spalten `Fehlt in B`, `Fehlt in A` und `Abweichend` je Verzeichnis.
* Abweichende Groß-/Kleinschreibung wird als Unterschied ausgewiesen, mit beiden
  Schreibweisen im Bericht.
* **Korrektur:** `-ExcludeFolder` hat auch auf Dateinamen gefiltert.
* **Korrektur:** Laufwerkstammverzeichnisse wie `D:\` wurden auf `D:` gekürzt, was
  unter Windows das *aktuelle* Verzeichnis dieses Laufwerks meint und damit den
  falschen Ordner einlesen konnte.

### Version 1.0

* Erste Fassung.
