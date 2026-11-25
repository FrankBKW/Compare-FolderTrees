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

### Parameter zum Kopieren

| Parameter          | Standard        | Zweck                                                     |
|--------------------|-----------------|-----------------------------------------------------------|
| `-Kopieren`        | *aus*           | `AnachB`, `BnachA` oder `Beide` — ohne diesen Parameter wird nichts geschrieben |
| `-KopiereWas`      | `Fehlende`      | `Fehlende`, `Unterschiedliche` oder `Alles`               |
| `-BeiKonflikt`     | `Ueberspringen` | bei `Beide` und abweichenden Dateien: `Ueberspringen` oder `Neuere` |
| `-Ausfuehren`      | aus             | wirklich kopieren — ohne diesen Schalter nur Simulation   |
| `-Sicherung`       | aus             | überschriebene Dateien vorher wegsichern                   |
| `-TrotzLesefehler` | aus             | kopieren, obwohl die Zielseite unvollständig gelesen wurde |
| `-Auswahl`         | *aus*           | Pfad zu einer im HTML-Bericht angekreuzten Auswahlliste    |

## Dateien im Bericht ankreuzen

Statt ganzer Kategorien lassen sich einzelne Dateien auswählen. In der
Dateitabelle stehen dafür zwei Ankreuzspalten ganz links:

| Spalte | Bedeutung                        |
|--------|----------------------------------|
| **→ B** | diese Datei von A nach B kopieren |
| **→ A** | diese Datei von B nach A kopieren |

Pro Zeile ist nur eine Richtung möglich — das zweite Häkchen schaltet das erste
ab. Angeboten wird nur, was auch geht: bei `Nur in A` gibt es ausschließlich
**→ B**, bei `Nur in B` nur **→ A**, bei `Unterschiedlich` beides. Identische
Dateien haben keine Ankreuzfelder.

Unter der Tabelle zeigt ein Kasten laufend mit, wie viele Dateien und wie viele
Megabyte ausgewählt sind, getrennt nach Richtung. Die Schaltflächen **Alle
sichtbaren → B** und **→ A** wirken nur auf die gerade gefilterten Zeilen —
damit lässt sich etwa über „Nur Abweichungen" vorfiltern und dann alles auf
einmal ankreuzen.

### Warum über eine Liste und nicht direkt

Eine lokal geöffnete HTML-Seite darf aus Sicherheitsgründen nicht auf das
Dateisystem schreiben; kein Browser erlaubt das. Der Bericht sammelt deshalb nur
die Häkchen ein und erzeugt daraus eine Auswahlliste. Kopiert wird
ausschließlich vom Script — mit denselben Schutzmaßnahmen wie sonst auch.

```powershell
# 1. Im Bericht ankreuzen, dann "Auswahlliste speichern" (landet als Auswahl.txt)

# 2. Ansehen, was passieren würde:
.\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup" -Auswahl "$HOME\Downloads\Auswahl.txt"

# 3. Ausführen:
.\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup" -Auswahl "$HOME\Downloads\Auswahl.txt" -Ausfuehren
```

Den fertigen Befehl zeigt der Bericht unter den Schaltflächen an. Falls der
Browser das Speichern blockiert, steht die Liste zusätzlich zum Kopieren in
einem aufklappbaren Textfeld.

Die Auswahlliste merkt sich, zu welchem Ordnerpaar sie gehört. Wird sie
versehentlich auf andere Ordner angewendet, bricht das Script ab und nennt
beide Pfadpaare. Einträge, die nicht mehr zum aktuellen Vergleich passen — etwa
weil der Bericht veraltet ist —, werden einzeln gemeldet und übersprungen,
statt irgendetwas zu raten.

## Kopieren

Fehlende oder abweichende Dateien lassen sich in den jeweils anderen Ordner
übertragen. Der Vergleich bleibt dabei unverändert — das Kopieren setzt nur auf
dessen Ergebnis auf.

```powershell
# 1. Ansehen, was passieren würde. Schreibt nichts.
.\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup" -Kopieren AnachB

# 2. Wenn der Plan stimmt: ausführen.
.\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup" -Kopieren AnachB -Ausfuehren
```

### Was passiert wann

`-Kopieren AnachB` überträgt von A nach B, `BnachA` umgekehrt. `Beide` füllt
fehlende Dateien in beide Richtungen auf.

`-KopiereWas Fehlende` (Standard) rührt vorhandene Dateien nicht an. Erst
`Unterschiedliche` oder `Alles` überschreibt auch bestehende Dateien.

Bei `-Kopieren Beide` zusammen mit abweichenden Dateien ist nicht von selbst
klar, welche Seite gewinnt. Solche Fälle werden deshalb standardmäßig
übersprungen und als Konflikt aufgeführt. Mit `-BeiKonflikt Neuere` entscheidet
das Änderungsdatum; sind beide Seiten gleich alt, aber inhaltlich verschieden,
bleibt die Datei unangetastet, weil keine Seite erkennbar aktueller ist.

### Schutzmaßnahmen

Kopieren verändert echte Datenbestände, deshalb mehrere Sicherungen:

* **Simulation ist der Standard.** Ohne `-Ausfuehren` entsteht nur ein Plan.
* **Es wird nie gelöscht.** Dateien, die auf der Zielseite zu viel sind, bleiben
  liegen — das Script ist kein Spiegel-Abgleich.
* **Zielprüfung unmittelbar vor dem Schreiben.** Liegt dort wider Erwarten schon
  eine Datei, obwohl sie als fehlend eingestuft war, wird sie nicht
  überschrieben, sondern der Fall im Bericht ausgewiesen.
* **Abbruch bei unvollständig gelesener Zielseite.** Konnte das Ziel nicht ganz
  eingelesen werden, ist „fehlt im Ziel" nicht verlässlich. Das Script kopiert
  dann nichts, bis die Ursache behoben ist oder `-TrotzLesefehler` gesetzt wird.
* **`-Sicherung`** legt überschriebene Dateien vorher unter
  `<OutputFolder>\Sicherung_<Zeitstempel>\` ab, mit erhaltener Ordnerstruktur.

Zeitstempel werden beim Kopieren übernommen. Ein zweiter Lauf findet die Dateien
deshalb als identisch vor und kopiert nichts erneut.

### Ergebnis

Der HTML-Bericht bekommt einen Abschnitt **3. Kopieren** mit allen Vorgängen,
getrennt nach neu angelegt, überschrieben, übersprungen und fehlgeschlagen.
Zusätzlich entsteht `Kopieren_<Zeitstempel>.csv` mit Quelle, Ziel, Aktion und
Ergebnis je Datei.

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

### Version 1.5

* **Neu: Dateien im Bericht ankreuzen.** Die Dateitabelle hat zwei
  Ankreuzspalten **→ B** und **→ A**, pro Zeile gegenseitig ausschließend und
  nur dort angeboten, wo die Richtung möglich ist. Ein Kasten unter der Tabelle
  zählt die Auswahl mit und erzeugt daraus eine Auswahlliste, die das Script
  über `-Auswahl` ausführt. Die Liste ist an ihr Ordnerpaar gebunden; auf
  andere Ordner angewendet bricht das Script ab.
* **Korrektur:** Mit `-Ausfuehren`, aber leerem Kopierplan blieben abgewiesene
  Einträge ohne Ergebnis stehen und die Zusammenfassung meldete „Simulation",
  obwohl ausgeführt wurde.

### Version 1.4

* **Neu: Kopieren.** Fehlende und abweichende Dateien lassen sich wahlweise in
  den anderen Ordner übertragen — `-Kopieren AnachB|BnachA|Beide`, wahlweise
  eingegrenzt mit `-KopiereWas`. Ohne `-Ausfuehren` wird nur simuliert, gelöscht
  wird nie, und vor jedem Schreibvorgang prüft das Script das Ziel noch einmal.
  Mit `-Sicherung` landen überschriebene Dateien vorher in einem Sicherungsordner.
  Neuer Abschnitt im HTML-Bericht und eine eigene `Kopieren_*.csv`.
* **Korrektur:** `Get-Help` zeigte nur die nackte Syntaxzeile statt der
  Dokumentation. Ursache war die `#Requires`-Zeile oberhalb des Hilfeblocks —
  damit erkennt Windows PowerShell 5.1 die kommentarbasierte Hilfe nicht. Sie
  steht jetzt darunter.

### Version 1.3.1

* **Korrektur:** Spaltenköpfe standen nicht über ihren Werten. Die Überschriften
  waren linksbündig, Größen und Anzahlen darunter rechtsbündig — Kopf und Wert
  saßen damit an entgegengesetzten Rändern derselben Spalte. Zahlenspalten sind
  jetzt durchgehend rechtsbündig, Datumsspalten durchgehend linksbündig.
* Dateinamen werden nicht mehr mitten im Wort umbrochen; die Namensspalte hat
  eine Mindestbreite, damit sie von den Zahlenspalten nicht zusammengedrückt wird.

### Version 1.3

* **Warnung bei unvollständigem Einlesen.** Konnte eine Seite nicht komplett
  gelesen werden, steht das jetzt als rot abgesetzter Kasten oben im Bericht und
  in der Konsolenausgabe — mit der Angabe, welche Seite betroffen ist und dass
  deren Dateien fälschlich als „nur auf der anderen Seite" erscheinen können.
  Bisher war das nur unten unter „Hinweise" zu finden und damit leicht zu
  übersehen. Das ist die häufigste Ursache für scheinbar falsche Ergebnisse:
  Pfade über 260 Zeichen treffen nur die tiefer verschachtelte Seite.

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
