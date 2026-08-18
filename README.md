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
zugeordnet (Groß-/Kleinschreibung wird ignoriert). Jedes Objekt bekommt einen von
vier Status:

| Status            | Bedeutung                                              |
|-------------------|--------------------------------------------------------|
| `Nur in A`        | existiert nur im ersten Ordner — fehlt in B            |
| `Nur in B`        | existiert nur im zweiten Ordner — fehlt in A           |
| `Unterschiedlich` | auf beiden Seiten vorhanden, aber abweichend           |
| `Identisch`       | Größe und Änderungsdatum (optional Hash) stimmen überein |

Bei Dateien wird zusätzlich ausgewiesen, *was* abweicht: Größe, Änderungsdatum
oder — mit `-CompareHash` — der Inhalt.

## Ausgabe

### 1. Verzeichnisse

Relativer Pfad, Größe A/B (Summe aller enthaltenen Dateien, rekursiv), Differenz,
Anzahl Dateien A/B, Hinweis.

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

| Datei                          | Inhalt                                          |
|--------------------------------|-------------------------------------------------|
| `Vergleich_Verzeichnisse_*.csv`| nur Verzeichnisse                               |
| `Vergleich_Dateien_*.csv`      | nur Dateien                                     |
| `Vergleich_Gesamt_*.csv`       | beides mit vereinheitlichten Spalten            |

Größen stehen jeweils als Rohwert in Bytes **und** formatiert (`12,34 MB`) in der
Tabelle, damit sich in Excel rechnen und sortieren lässt, ohne die Lesbarkeit zu
verlieren.

## Parameter

| Parameter                | Standard                        | Zweck                                                     |
|--------------------------|---------------------------------|-----------------------------------------------------------|
| `-PfadA`                 | *erforderlich*                  | erster Ordner (Referenz)                                   |
| `-PfadB`                 | *erforderlich*                  | zweiter Ordner (Vergleich)                                 |
| `-OutputFolder`          | `<Desktop>\Ordnervergleich`     | Zielordner für HTML und CSV                                |
| `-Filter`                | `*`                             | Dateifilter, z. B. `*.pdf`                                 |
| `-ExcludeFolder`         | *leer*                          | Ordnernamen überspringen, Wildcards erlaubt                |
| `-TimeToleranceSeconds`  | `2`                             | Toleranz beim Datumsvergleich (FAT/NTFS-Rundung)           |
| `-CompareHash`           | aus                             | zusätzlich Prüfsumme, nur bei gleicher Größe               |
| `-HashAlgorithm`         | `MD5`                           | `MD5`, `SHA1` oder `SHA256`                                |
| `-CsvDelimiter`          | `;`                             | auf `,` setzen für englischsprachiges Excel                |
| `-OpenReport`            | aus                             | HTML-Bericht am Ende öffnen                                |

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
