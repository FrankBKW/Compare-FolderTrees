#Requires -Version 5.1
<#
.SYNOPSIS
    Vergleicht zwei Ordnerstrukturen (inklusive aller Unterordner) und erzeugt
    einen HTML-Bericht sowie CSV-Dateien.

.DESCRIPTION
    Ausgabe je Verzeichnis : relativer Pfad, Groesse (Summe aller Dateien, rekursiv),
                             Anzahl Dateien - jeweils fuer Seite A und Seite B.
    Ausgabe je Datei       : Dateiname, Groesse und Aenderungsdatum - je Seite A und B.
    Fehlende Objekte werden als "Nur in A" bzw. "Nur in B" gekennzeichnet.

.PARAMETER PfadA
    Erster Ordner (Referenz).

.PARAMETER PfadB
    Zweiter Ordner (Vergleich).

.PARAMETER OutputFolder
    Zielordner fuer HTML- und CSV-Dateien. Standard: <Desktop>\Ordnervergleich

.PARAMETER Filter
    Dateifilter, z.B. '*.pdf'. Standard: '*'

.PARAMETER ExcludeFolder
    Ordnernamen, die uebersprungen werden (Wildcards erlaubt), z.B. '.git','node_modules'

.PARAMETER TimeToleranceSeconds
    Toleranz beim Vergleich des Aenderungsdatums (Standard 2 Sekunden, gleicht
    FAT/NTFS-Rundungen aus).

.PARAMETER CompareHash
    Zusaetzlich Pruefsumme vergleichen (nur wenn die Groesse gleich ist).
    Deutlich langsamer bei vielen/grossen Dateien.

.PARAMETER HashAlgorithm
    MD5 (Standard), SHA1 oder SHA256.

.PARAMETER CsvDelimiter
    Trennzeichen der CSV-Dateien. Standard ';' (Excel deutsch).

.PARAMETER OpenReport
    HTML-Report am Ende automatisch oeffnen.

.EXAMPLE
    .\Compare-FolderTrees.ps1 -PfadA "D:\Daten" -PfadB "E:\Backup\Daten" -OpenReport

.EXAMPLE
    .\Compare-FolderTrees.ps1 "D:\Projekte" "\\nas\projekte" -ExcludeFolder '.git','node_modules' -CompareHash

.NOTES
    Version 1.3.1

    Der Status eines Verzeichnisses wird aus den enthaltenen Dateien abgeleitet,
    nicht aus Gesamtgroesse und Dateianzahl: zwei Ordner koennen zufaellig gleich
    gross sein und trotzdem verschiedene Dateien enthalten.

    Dateinamen werden vor dem Vergleich unicode-normalisiert (Form C). Umlaute
    lassen sich auf zwei Arten speichern - als ein Zeichen oder als Grundzeichen
    plus kombinierendem Trema. Ohne Normalisierung erscheint dieselbe Datei
    zweimal im Bericht: einmal als "Nur in A" und einmal als "Nur in B".
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$PfadA,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$PfadB,

    [string]$OutputFolder = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Ordnervergleich'),

    [string]$Filter = '*',

    [string[]]$ExcludeFolder = @(),

    [int]$TimeToleranceSeconds = 2,

    [switch]$CompareHash,

    [ValidateSet('MD5', 'SHA1', 'SHA256')]
    [string]$HashAlgorithm = 'MD5',

    [string]$CsvDelimiter = ';',

    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'
$startTime = Get-Date

# ----------------------------------------------------------------------------
# Hilfsfunktionen
# ----------------------------------------------------------------------------

function Format-Size {
    param($Bytes = $null)
    if ($null -eq $Bytes) { return '' }
    $v = [double]$Bytes
    if ($v -lt 0) { return '' }
    $units = @('B', 'KB', 'MB', 'GB', 'TB', 'PB')
    $i = 0
    while ($v -ge 1024 -and $i -lt ($units.Count - 1)) {
        $v = $v / 1024
        $i++
    }
    if ($i -eq 0) { return ('{0:N0} {1}' -f $v, $units[$i]) }
    return ('{0:N2} {1}' -f $v, $units[$i])
}

function ConvertTo-HtmlText {
    param($Text = '')
    if ($null -eq $Text) { return '' }
    $t = [string]$Text
    if ($t.Length -eq 0) { return '' }
    $t = $t -replace '&', '&amp;'
    $t = $t -replace '<', '&lt;'
    $t = $t -replace '>', '&gt;'
    $t = $t -replace '"', '&quot;'
    return $t
}

function Format-Date {
    param($Value = $null)
    if ($null -eq $Value) { return '' }
    return ([datetime]$Value).ToString('yyyy-MM-dd HH:mm:ss')
}

function ConvertTo-NormKey {
    # Unicode-Normalisierung auf Form C.
    #
    # Umlaute koennen auf zwei Arten gespeichert sein: "ue" als ein Zeichen
    # (U+00FC, so schreibt Windows) oder als "u" plus kombinierendes Trema
    # (U+0075 U+0308, so schreiben macOS und viele NAS-Systeme). Beide sehen
    # identisch aus, sind als Zeichenkette aber verschieden. Ohne diese
    # Normalisierung erscheint dieselbe Datei zweimal im Bericht - einmal als
    # "Nur in A" und einmal als "Nur in B".
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return $Text }
    try   { return $Text.Normalize([System.Text.NormalizationForm]::FormC) }
    catch { return $Text }   # ungueltige Zeichenfolgen unveraendert lassen
}

function Get-ParentRel {
    param([string]$Rel)
    $i = $Rel.LastIndexOf('\')
    if ($i -lt 0) { return '.' }
    return $Rel.Substring(0, $i)
}

function Get-LeafRel {
    param([string]$Rel)
    $i = $Rel.LastIndexOf('\')
    if ($i -lt 0) { return $Rel }
    return $Rel.Substring($i + 1)
}

function Get-NameDifference {
    # Liefert die Gruende, warum zwei Namen abweichen, die derselben Datei bzw.
    # demselben Ordner zugeordnet wurden. Leer, wenn sie exakt gleich sind.
    param([string]$NameA, [string]$NameB)

    $gruende = New-Object 'System.Collections.Generic.List[string]'
    if ([string]::Equals($NameA, $NameB, [StringComparison]::Ordinal)) { return $gruende }

    $nA = ConvertTo-NormKey $NameA
    $nB = ConvertTo-NormKey $NameB

    # Nach der Normalisierung noch unterschiedlich -> Gross-/Kleinschreibung
    if (-not [string]::Equals($nA, $nB, [StringComparison]::Ordinal)) { $gruende.Add('Schreibweise') }
    # Roh unterschiedlich, obwohl normalisiert gleich -> abweichende Kodierung
    if (-not [string]::Equals($NameA, $NameB, [StringComparison]::OrdinalIgnoreCase)) { $gruende.Add('Namenskodierung') }

    return $gruende
}

function Get-StatusKey {
    param([string]$Status)
    switch ($Status) {
        'Nur in A'        { return 'onlyA' }
        'Nur in B'        { return 'onlyB' }
        'Unterschiedlich' { return 'diff' }
        default           { return 'same' }
    }
}

function Get-FileHashSafe {
    param([string]$Path, [string]$Algorithm)
    try {
        return (Get-FileHash -LiteralPath $Path -Algorithm $Algorithm -ErrorAction Stop).Hash
    } catch {
        return ''
    }
}

# ----------------------------------------------------------------------------
# Verzeichnisbaum einlesen
# ----------------------------------------------------------------------------

function Get-TreeIndex {
    param(
        [string]$Root,
        [string]$Filter,
        [string[]]$ExcludeFolder,
        [string]$Label
    )

    $rootFull = (Resolve-Path -LiteralPath $Root).ProviderPath.TrimEnd('\')
    # "D:" bedeutet unter Windows "aktuelles Verzeichnis auf D:", nicht das
    # Laufwerkstammverzeichnis - deshalb den Backslash wieder anhaengen.
    if ($rootFull -match '^[A-Za-z]:$') { $rootFull = $rootFull + '\' }
    $prefixLen = $rootFull.Length

    $cmp    = [StringComparer]::OrdinalIgnoreCase
    $files  = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList $cmp
    $dirs   = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList $cmp
    $errors = New-Object 'System.Collections.Generic.List[string]'

    Write-Host ("[{0}] Lese '{1}' ..." -f $Label, $rootFull) -ForegroundColor Cyan

    $scanErr = $null
    $items = Get-ChildItem -LiteralPath $rootFull -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable scanErr
    if ($scanErr) {
        foreach ($e in $scanErr) { $errors.Add("[$Label] " + [string]$e) }
    }

    # Wurzelverzeichnis als eigener Eintrag
    $rootItem = Get-Item -LiteralPath $rootFull -Force
    $dirs['.'] = [pscustomobject]@{
        Rel       = '.'
        Name      = '.'
        Full      = $rootFull
        Size      = [long]0
        FileCount = 0
        LastWrite = $rootItem.LastWriteTime
    }

    foreach ($item in $items) {
        # Der relative Pfad dient als Schluessel und wird dafuer normalisiert.
        # Die tatsaechliche Schreibweise bleibt in Name/Full erhalten.
        $rel = ConvertTo-NormKey ($item.FullName.Substring($prefixLen).TrimStart('\'))
        if ([string]::IsNullOrEmpty($rel)) { continue }

        if ($ExcludeFolder.Count -gt 0) {
            # Nur Ordnersegmente pruefen - bei Dateien bleibt der Dateiname aussen vor,
            # sonst wuerde -ExcludeFolder ungewollt auch Dateinamen filtern.
            $segs = $rel.Split('\')
            $anzahl = $segs.Count
            if (-not $item.PSIsContainer) { $anzahl = $anzahl - 1 }
            $skip = $false
            for ($s = 0; $s -lt $anzahl; $s++) {
                foreach ($ex in $ExcludeFolder) {
                    if ($segs[$s] -like $ex) { $skip = $true; break }
                }
                if ($skip) { break }
            }
            if ($skip) { continue }
        }

        if ($item.PSIsContainer) {
            # Junctions und Ordner-Symlinks werden von Get-ChildItem -Recurse nicht
            # verfolgt. Der Inhalt fehlt dadurch im Vergleich, obwohl er im Explorer
            # sichtbar ist - deshalb ausdruecklich vermerken statt still uebergehen.
            if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
                $errors.Add("[$Label] Verknuepfter Ordner wird nicht verfolgt, Inhalt fehlt im Vergleich: " + $item.FullName)
            }
            if (-not $dirs.ContainsKey($rel)) {
                $dirs[$rel] = [pscustomobject]@{
                    Rel       = $rel
                    Name      = $item.Name
                    Full      = $item.FullName
                    Size      = [long]0
                    FileCount = 0
                    LastWrite = $item.LastWriteTime
                }
            }
        } else {
            if ($Filter -ne '*' -and $item.Name -notlike $Filter) { continue }

            $files[$rel] = [pscustomobject]@{
                Rel       = $rel
                Full      = $item.FullName
                Name      = $item.Name
                Dir       = (Get-ParentRel $rel)
                Size      = [long]$item.Length
                LastWrite = $item.LastWriteTime
            }
        }
    }

    # Verzeichnisgroessen rekursiv aufsummieren
    # (jede Datei zaehlt fuer ihren Ordner und alle uebergeordneten Ordner)
    foreach ($f in $files.Values) {
        $parent = $f.Dir
        while ($true) {
            if (-not $dirs.ContainsKey($parent)) {
                $dirs[$parent] = [pscustomobject]@{
                    Rel       = $parent
                    Name      = (Get-LeafRel $parent)
                    Full      = (Join-Path $rootFull $parent)
                    Size      = [long]0
                    FileCount = 0
                    LastWrite = $null
                }
            }
            $dirs[$parent].Size      = $dirs[$parent].Size + $f.Size
            $dirs[$parent].FileCount = $dirs[$parent].FileCount + 1

            if ($parent -eq '.') { break }
            $parent = Get-ParentRel $parent
        }
    }

    Write-Host ("[{0}] {1} Ordner, {2} Dateien" -f $Label, ($dirs.Count - 1), $files.Count) -ForegroundColor DarkGray

    return [pscustomobject]@{
        Root   = $rootFull
        Files  = $files
        Dirs   = $dirs
        Errors = $errors
    }
}

# ----------------------------------------------------------------------------
# Vorbereitung
# ----------------------------------------------------------------------------

if (-not (Test-Path -LiteralPath $PfadA)) { throw "Pfad A existiert nicht: $PfadA" }
if (-not (Test-Path -LiteralPath $PfadB)) { throw "Pfad B existiert nicht: $PfadB" }

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}
$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath

$idxA = Get-TreeIndex -Root $PfadA -Filter $Filter -ExcludeFolder $ExcludeFolder -Label 'A'
$idxB = Get-TreeIndex -Root $PfadB -Filter $Filter -ExcludeFolder $ExcludeFolder -Label 'B'

if ($idxA.Root -eq $idxB.Root) { throw 'Pfad A und Pfad B verweisen auf denselben Ordner.' }

# ----------------------------------------------------------------------------
# Dateien vergleichen
#
# Wichtig: Die Dateien werden VOR den Verzeichnissen ausgewertet, weil der
# Status eines Verzeichnisses aus den enthaltenen Dateien abgeleitet wird.
# ----------------------------------------------------------------------------

Write-Host 'Vergleiche Dateien ...' -ForegroundColor Cyan

$fileKeys = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
foreach ($k in $idxA.Files.Keys) { [void]$fileKeys.Add($k) }
foreach ($k in $idxB.Files.Keys) { [void]$fileKeys.Add($k) }

$fileRows = New-Object 'System.Collections.Generic.List[object]'
$total = $fileKeys.Count
$n = 0

foreach ($key in ($fileKeys | Sort-Object)) {
    $n++
    if ($total -gt 0 -and ($n % 250 -eq 0 -or $n -eq $total)) {
        Write-Progress -Activity 'Dateivergleich' -Status "$n / $total" -PercentComplete (($n / $total) * 100)
    }

    $a = $null; $b = $null
    $hasA = $idxA.Files.TryGetValue($key, [ref]$a)
    $hasB = $idxB.Files.TryGetValue($key, [ref]$b)

    $unterschied = ''
    $status      = 'Identisch'
    $hashA       = ''
    $hashB       = ''

    if ($hasA -and -not $hasB) {
        $status = 'Nur in A'; $unterschied = 'Fehlt in B'
    } elseif ($hasB -and -not $hasA) {
        $status = 'Nur in B'; $unterschied = 'Fehlt in A'
    } else {
        $r = New-Object 'System.Collections.Generic.List[string]'
        if ($a.Size -ne $b.Size) { $r.Add('Groesse') }

        $deltaSec = [math]::Abs(($b.LastWrite - $a.LastWrite).TotalSeconds)
        if ($deltaSec -gt $TimeToleranceSeconds) { $r.Add('Aenderungsdatum') }

        # Namen weichen ab, obwohl sie auf denselben Schluessel fuehren: entweder
        # in der Gross-/Kleinschreibung oder in der Unicode-Kodierung der Umlaute.
        # Verglichen wird nur der Dateiname, damit ein abweichend geschriebener
        # Ordnername nicht jede Datei darunter als abweichend markiert.
        foreach ($grund in (Get-NameDifference -NameA $a.Name -NameB $b.Name)) { $r.Add($grund) }

        if ($CompareHash -and $a.Size -eq $b.Size) {
            $hashA = Get-FileHashSafe -Path $a.Full -Algorithm $HashAlgorithm
            $hashB = Get-FileHashSafe -Path $b.Full -Algorithm $HashAlgorithm
            if ($hashA -and $hashB -and $hashA -ne $hashB) { $r.Add('Inhalt (Hash)') }
        }

        if ($r.Count -gt 0) { $status = 'Unterschiedlich'; $unterschied = ($r -join ', ') }
    }

    $sizeA = $null; $dateA = ''; $fullA = ''
    $sizeB = $null; $dateB = ''; $fullB = ''
    $diff  = $null
    if ($hasA) { $sizeA = [long]$a.Size; $dateA = Format-Date $a.LastWrite; $fullA = $a.Full }
    if ($hasB) { $sizeB = [long]$b.Size; $dateB = Format-Date $b.LastWrite; $fullB = $b.Full }
    if ($hasA -and $hasB) { $diff = [long]($b.Size - $a.Size) }

    $verzeichnis = ''
    $dateiname   = ''
    if ($hasA) { $verzeichnis = $a.Dir; $dateiname = $a.Name }
    else       { $verzeichnis = $b.Dir; $dateiname = $b.Name }

    # Bei abweichender Schreibweise beide Varianten ausweisen
    $dateinameA = ''
    $dateinameB = ''
    if ($hasA) { $dateinameA = $a.Name }
    if ($hasB) { $dateinameB = $b.Name }

    $fileRows.Add([pscustomobject]@{
        Typ             = 'Datei'
        Status          = $status
        RelativerPfad   = $key
        Dateiname       = $dateiname
        DateinameA      = $dateinameA
        DateinameB      = $dateinameB
        Verzeichnis     = $verzeichnis
        GroesseA_Bytes  = $sizeA
        GroesseB_Bytes  = $sizeB
        GroesseA        = (Format-Size $sizeA)
        GroesseB        = (Format-Size $sizeB)
        Differenz_Bytes = $diff
        GeaendertA      = $dateA
        GeaendertB      = $dateB
        Unterschied     = $unterschied
        HashA           = $hashA
        HashB           = $hashB
        PfadA           = $fullA
        PfadB           = $fullB
    })
}
Write-Progress -Activity 'Dateivergleich' -Completed

# ----------------------------------------------------------------------------
# Verzeichnisse vergleichen
#
# Der Status ergibt sich aus den enthaltenen Dateien (rekursiv), nicht aus
# Groesse und Anzahl: zwei Ordner koennen zufaellig gleich gross sein und
# trotzdem voellig verschiedene Dateien enthalten.
# ----------------------------------------------------------------------------

Write-Host 'Vergleiche Verzeichnisse ...' -ForegroundColor Cyan

# Abweichungen der Dateien auf alle uebergeordneten Ordner hochrechnen
$dirMarks = New-Object 'System.Collections.Generic.Dictionary[string,object]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)

function Add-DirMark {
    param([string]$StartDir, [string]$Art)
    $p = $StartDir
    while ($true) {
        $m = $null
        if (-not $dirMarks.TryGetValue($p, [ref]$m)) {
            $m = [pscustomobject]@{ FehltInB = 0; FehltInA = 0; Abweichend = 0 }
            $dirMarks[$p] = $m
        }
        switch ($Art) {
            'FehltInB'   { $m.FehltInB   = $m.FehltInB + 1 }
            'FehltInA'   { $m.FehltInA   = $m.FehltInA + 1 }
            'Abweichend' { $m.Abweichend = $m.Abweichend + 1 }
        }
        if ($p -eq '.') { break }
        $p = Get-ParentRel $p
    }
}

foreach ($f in $fileRows) {
    switch ($f.Status) {
        'Nur in A'        { Add-DirMark -StartDir $f.Verzeichnis -Art 'FehltInB' }
        'Nur in B'        { Add-DirMark -StartDir $f.Verzeichnis -Art 'FehltInA' }
        'Unterschiedlich' { Add-DirMark -StartDir $f.Verzeichnis -Art 'Abweichend' }
    }
}

$dirKeys = New-Object 'System.Collections.Generic.HashSet[string]' -ArgumentList ([StringComparer]::OrdinalIgnoreCase)
foreach ($k in $idxA.Dirs.Keys) { [void]$dirKeys.Add($k) }
foreach ($k in $idxB.Dirs.Keys) { [void]$dirKeys.Add($k) }

$dirRows = New-Object 'System.Collections.Generic.List[object]'

foreach ($key in ($dirKeys | Sort-Object)) {
    $a = $null; $b = $null
    $hasA = $idxA.Dirs.TryGetValue($key, [ref]$a)
    $hasB = $idxB.Dirs.TryGetValue($key, [ref]$b)

    $mark = $null
    $fehltInB = 0; $fehltInA = 0; $abweichend = 0
    if ($dirMarks.TryGetValue($key, [ref]$mark)) {
        $fehltInB = $mark.FehltInB; $fehltInA = $mark.FehltInA; $abweichend = $mark.Abweichend
    }

    $hinweis = ''
    $status  = 'Identisch'
    if ($hasA -and -not $hasB) {
        $status = 'Nur in A'; $hinweis = 'Ordner fehlt in B'
    } elseif ($hasB -and -not $hasA) {
        $status = 'Nur in B'; $hinweis = 'Ordner fehlt in A'
    } else {
        $r = New-Object 'System.Collections.Generic.List[string]'
        if ($fehltInB -gt 0) {
            $verb = 'fehlen'; if ($fehltInB -eq 1) { $verb = 'fehlt' }
            $r.Add("$fehltInB $verb in B")
        }
        if ($fehltInA -gt 0) {
            $verb = 'fehlen'; if ($fehltInA -eq 1) { $verb = 'fehlt' }
            $r.Add("$fehltInA $verb in A")
        }
        if ($abweichend -gt 0) { $r.Add("$abweichend abweichend") }
        # Name des Ordners selbst (nur die letzte Ebene, damit ein abweichend
        # geschriebener Elternordner nicht den ganzen Teilbaum markiert)
        if ($key -ne '.') {
            foreach ($grund in (Get-NameDifference -NameA $a.Name -NameB $b.Name)) { $r.Add($grund) }
        }
        if ($r.Count -gt 0) { $status = 'Unterschiedlich'; $hinweis = ($r -join ', ') }
    }

    $sizeA  = $null; $countA = $null; $fullA = ''
    $sizeB  = $null; $countB = $null; $fullB = ''
    $diff   = $null
    if ($hasA) { $sizeA = [long]$a.Size; $countA = [int]$a.FileCount; $fullA = $a.Full }
    if ($hasB) { $sizeB = [long]$b.Size; $countB = [int]$b.FileCount; $fullB = $b.Full }
    if ($hasA -and $hasB) { $diff = [long]($b.Size - $a.Size) }

    $name  = '(Wurzelverzeichnis)'
    $ebene = 0
    if ($key -ne '.') {
        $name  = Get-LeafRel $key
        $ebene = $key.Split('\').Count
    }

    $dirRows.Add([pscustomobject]@{
        Typ             = 'Verzeichnis'
        Status          = $status
        RelativerPfad   = $key
        Name            = $name
        Ebene           = $ebene
        GroesseA_Bytes  = $sizeA
        GroesseB_Bytes  = $sizeB
        GroesseA        = (Format-Size $sizeA)
        GroesseB        = (Format-Size $sizeB)
        Differenz_Bytes = $diff
        DateienA        = $countA
        DateienB        = $countB
        FehltInB        = $fehltInB
        FehltInA        = $fehltInA
        Abweichend      = $abweichend
        Hinweis         = $hinweis
        PfadA           = $fullA
        PfadB           = $fullB
    })
}

# ----------------------------------------------------------------------------
# Kennzahlen
# ----------------------------------------------------------------------------

$stamp = Get-Date -Format 'yyyy-MM-dd_HHmmss'

$fOnlyA = @($fileRows | Where-Object { $_.Status -eq 'Nur in A' }).Count
$fOnlyB = @($fileRows | Where-Object { $_.Status -eq 'Nur in B' }).Count
$fDiff  = @($fileRows | Where-Object { $_.Status -eq 'Unterschiedlich' }).Count
$fSame  = @($fileRows | Where-Object { $_.Status -eq 'Identisch' }).Count

$dOnlyA = @($dirRows | Where-Object { $_.Status -eq 'Nur in A' }).Count
$dOnlyB = @($dirRows | Where-Object { $_.Status -eq 'Nur in B' }).Count
$dDiff  = @($dirRows | Where-Object { $_.Status -eq 'Unterschiedlich' }).Count
$dSame  = @($dirRows | Where-Object { $_.Status -eq 'Identisch' }).Count

$totalA = [long]0; foreach ($f in $idxA.Files.Values) { $totalA += $f.Size }
$totalB = [long]0; foreach ($f in $idxB.Files.Values) { $totalB += $f.Size }

# ----------------------------------------------------------------------------
# CSV-Ausgabe
# ----------------------------------------------------------------------------

Write-Host 'Schreibe CSV ...' -ForegroundColor Cyan

$csvDirs  = Join-Path $OutputFolder "Vergleich_Verzeichnisse_$stamp.csv"
$csvFiles = Join-Path $OutputFolder "Vergleich_Dateien_$stamp.csv"
$csvAll   = Join-Path $OutputFolder "Vergleich_Gesamt_$stamp.csv"

$dirRows  | Export-Csv -LiteralPath $csvDirs  -Delimiter $CsvDelimiter -Encoding UTF8 -NoTypeInformation
$fileRows | Export-Csv -LiteralPath $csvFiles -Delimiter $CsvDelimiter -Encoding UTF8 -NoTypeInformation

# Gemeinsame CSV mit einheitlichen Spalten (Verzeichnisse + Dateien)
$combined = New-Object 'System.Collections.Generic.List[object]'
foreach ($d in $dirRows) {
    $combined.Add([pscustomobject]@{
        Typ             = $d.Typ
        Status          = $d.Status
        RelativerPfad   = $d.RelativerPfad
        Name            = $d.Name
        GroesseA_Bytes  = $d.GroesseA_Bytes
        GroesseB_Bytes  = $d.GroesseB_Bytes
        GroesseA        = $d.GroesseA
        GroesseB        = $d.GroesseB
        Differenz_Bytes = $d.Differenz_Bytes
        GeaendertA      = ''
        GeaendertB      = ''
        DateienA        = $d.DateienA
        DateienB        = $d.DateienB
        FehltInB        = $d.FehltInB
        FehltInA        = $d.FehltInA
        Abweichend      = $d.Abweichend
        Unterschied     = $d.Hinweis
        PfadA           = $d.PfadA
        PfadB           = $d.PfadB
    })
}
foreach ($f in $fileRows) {
    $combined.Add([pscustomobject]@{
        Typ             = $f.Typ
        Status          = $f.Status
        RelativerPfad   = $f.RelativerPfad
        Name            = $f.Dateiname
        GroesseA_Bytes  = $f.GroesseA_Bytes
        GroesseB_Bytes  = $f.GroesseB_Bytes
        GroesseA        = $f.GroesseA
        GroesseB        = $f.GroesseB
        Differenz_Bytes = $f.Differenz_Bytes
        GeaendertA      = $f.GeaendertA
        GeaendertB      = $f.GeaendertB
        DateienA        = ''
        DateienB        = ''
        FehltInB        = ''
        FehltInA        = ''
        Abweichend      = ''
        Unterschied     = $f.Unterschied
        PfadA           = $f.PfadA
        PfadB           = $f.PfadB
    })
}
$combined | Export-Csv -LiteralPath $csvAll -Delimiter $CsvDelimiter -Encoding UTF8 -NoTypeInformation

# ----------------------------------------------------------------------------
# HTML-Ausgabe
# ----------------------------------------------------------------------------

Write-Host 'Erzeuge HTML-Report ...' -ForegroundColor Cyan

$css = @'
*{box-sizing:border-box}
body{margin:0;padding:24px;font-family:"Segoe UI",Roboto,Helvetica,Arial,sans-serif;
     background:#f4f6f8;color:#1f2933;font-size:14px}
h1{font-size:22px;margin:0 0 4px}
h2{font-size:17px;margin:32px 0 10px}
.sub{color:#6b7280;margin-bottom:20px}
.paths{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px}
.pathbox{flex:1 1 320px;background:#fff;border:1px solid #e1e5ea;border-radius:8px;padding:12px 14px}
.tag{display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;
     border-radius:5px;color:#fff;font-weight:700;margin-right:8px}
.tagA{background:#2563eb}.tagB{background:#7c3aed}
code{font-family:Consolas,"Courier New",monospace;word-break:break-all}
.cards{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:8px}
.card{background:#fff;border:1px solid #e1e5ea;border-radius:8px;padding:12px 16px;min-width:132px}
.card .k{font-size:11.5px;color:#6b7280;text-transform:uppercase;letter-spacing:.04em}
.card .v{font-size:22px;font-weight:600;margin-top:2px}
.card.onlyA .v{color:#2563eb}.card.onlyB .v{color:#7c3aed}
.card.diff .v{color:#d97706}.card.same .v{color:#059669}
.panel,.notes{background:#fff;border:1px solid #e1e5ea;border-radius:8px;overflow:hidden}
.toolbar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;padding:10px 12px;
     border-bottom:1px solid #e1e5ea;background:#fafbfc}
.toolbar input{flex:1 1 220px;padding:6px 10px;border:1px solid #cbd2d9;border-radius:6px;font-size:13px}
.filters{display:flex;gap:6px;flex-wrap:wrap}
.filters button{padding:6px 11px;border:1px solid #cbd2d9;background:#fff;border-radius:6px;
     cursor:pointer;font-size:12.5px}
.filters button:hover{background:#eef2f6}
.filters button.active{background:#1f2933;color:#fff;border-color:#1f2933}
.info{color:#6b7280;font-size:12.5px;margin-left:auto;white-space:nowrap}
.scroll{max-height:70vh;overflow:auto}
table{border-collapse:collapse;width:100%}
th,td{padding:7px 10px;border-bottom:1px solid #eef1f4;text-align:left;vertical-align:top}
th{position:sticky;top:0;background:#f0f3f6;cursor:pointer;user-select:none;white-space:nowrap;
   font-size:12px;text-transform:uppercase;letter-spacing:.03em;color:#52606d;z-index:1}
th:hover{background:#e4e9ee}
/* Zahlenspalten rechtsbuendig - der Spaltenkopf muss mitziehen, sonst stehen
   Ueberschrift und Wert an entgegengesetzten Raendern derselben Spalte. */
td.num{text-align:right;white-space:nowrap;font-variant-numeric:tabular-nums}
th.num{text-align:right}
/* Datumsspalten: gleiche Breite, aber linksbuendig wie ihr Kopf */
td.dt{text-align:left;white-space:nowrap;font-variant-numeric:tabular-nums}
/* Umbruch nur, wenn es sonst nicht passt - "break-all" zerlegt sonst jeden
   Dateinamen mitten im Wort. min-width verhindert, dass die Namensspalte von
   den vielen Zahlenspalten zusammengequetscht wird. */
td.path{font-family:Consolas,"Courier New",monospace;font-size:12.5px;
     word-break:normal;overflow-wrap:anywhere;min-width:11rem}
tbody tr:hover{background:#f7fafc}
.badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11.5px;
     font-weight:600;white-space:nowrap}
.b-onlyA{background:#dbeafe;color:#1d4ed8}
.b-onlyB{background:#ede9fe;color:#6d28d9}
.b-diff{background:#fef3c7;color:#b45309}
.b-same{background:#d1fae5;color:#047857}
tr[data-status="onlyA"] td:first-child{border-left:3px solid #2563eb}
tr[data-status="onlyB"] td:first-child{border-left:3px solid #7c3aed}
tr[data-status="diff"]  td:first-child{border-left:3px solid #d97706}
tr[data-status="same"]  td:first-child{border-left:3px solid #10b981}
.warn{background:#fef2f2;border:1px solid #fca5a5;border-left:5px solid #dc2626;
     border-radius:8px;padding:14px 16px;margin-bottom:18px}
.warn h3{margin:0 0 6px;font-size:15px;color:#991b1b}
.warn p{margin:0 0 6px}
.warn ul{margin:6px 0 0 18px;padding:0}
.missing{color:#b91c1c;font-weight:600}
.neg{color:#b91c1c}.pos{color:#047857}
footer{margin-top:28px;color:#6b7280;font-size:12.5px}
@media print{.toolbar{display:none}.scroll{max-height:none}}
'@

$js = @'
function initPanel(panel){
  var input = panel.querySelector('.q');
  // Ohne Suchfeld und Filter ist nichts zu verdrahten. Ohne diese Pruefung
  // bricht die Schleife beim ersten solchen Element ab - alle danach folgenden
  // Tabellen haetten dann keine funktionierenden Filterbuttons mehr.
  if (!input || !panel.querySelector('.filters button.active')) { return; }
  var buttons = panel.querySelectorAll('.filters button');
  for (var i = 0; i < buttons.length; i++){
    buttons[i].addEventListener('click', function(){
      var bs = panel.querySelectorAll('.filters button');
      for (var j = 0; j < bs.length; j++){ bs[j].classList.remove('active'); }
      this.classList.add('active');
      applyFilter(panel);
    });
  }
  input.addEventListener('input', function(){ applyFilter(panel); });
  var ths = panel.querySelectorAll('th');
  for (var k = 0; k < ths.length; k++){
    ths[k].addEventListener('click', function(){ sortTable(this); });
  }
  applyFilter(panel);
}

function applyFilter(panel){
  var q = panel.querySelector('.q').value.toLowerCase();
  var act = panel.querySelector('.filters button.active').getAttribute('data-status');
  var rows = panel.querySelectorAll('tbody tr');
  var shown = 0;
  for (var i = 0; i < rows.length; i++){
    var r = rows[i];
    var st = r.getAttribute('data-status');
    var okS = (act === 'all') || (st === act) || (act === 'notsame' && st !== 'same');
    var okQ = (q === '') || (r.getAttribute('data-search').indexOf(q) > -1);
    if (okS && okQ){ r.style.display = ''; shown++; } else { r.style.display = 'none'; }
  }
  panel.querySelector('.count').textContent = shown + ' von ' + rows.length + ' Zeilen';
}

function sortTable(th){
  var headRow = th.parentNode;
  var table = headRow.parentNode.parentNode;
  var tbody = table.tBodies[0];
  var idx = Array.prototype.indexOf.call(headRow.children, th);
  var asc = th.getAttribute('data-dir') !== 'asc';
  for (var h = 0; h < headRow.children.length; h++){
    headRow.children[h].removeAttribute('data-dir');
    headRow.children[h].textContent = headRow.children[h].textContent.replace(/ [\u25B2\u25BC]$/, '');
  }
  th.setAttribute('data-dir', asc ? 'asc' : 'desc');
  th.textContent = th.textContent + (asc ? ' \u25B2' : ' \u25BC');

  var rows = Array.prototype.slice.call(tbody.rows);
  rows.sort(function(a, b){
    var ca = a.cells[idx], cb = b.cells[idx];
    var va = ca.getAttribute('data-sort'), vb = cb.getAttribute('data-sort');
    if (va !== null && vb !== null){
      var na = parseFloat(va), nb = parseFloat(vb);
      if (!isNaN(na) && !isNaN(nb)){ return asc ? na - nb : nb - na; }
    }
    var ta = (va !== null ? va : ca.textContent).toLowerCase();
    var tb = (vb !== null ? vb : cb.textContent).toLowerCase();
    if (ta < tb) return asc ? -1 : 1;
    if (ta > tb) return asc ? 1 : -1;
    return 0;
  });
  for (var i = 0; i < rows.length; i++){ tbody.appendChild(rows[i]); }
}

document.addEventListener('DOMContentLoaded', function(){
  var panels = document.querySelectorAll('.panel');
  for (var i = 0; i < panels.length; i++){ initPanel(panels[i]); }
});
'@

function New-Toolbar {
    param([int]$OnlyA, [int]$OnlyB, [int]$Diff, [int]$Same)
    $h  = '<div class="toolbar">'
    $h += '<input class="q" type="search" placeholder="Suchen (Pfad / Name) ...">'
    $h += '<div class="filters">'
    $h += '<button data-status="all" class="active">Alle</button>'
    $h += '<button data-status="notsame">Nur Abweichungen</button>'
    $h += '<button data-status="onlyA">Nur in A (' + $OnlyA + ')</button>'
    $h += '<button data-status="onlyB">Nur in B (' + $OnlyB + ')</button>'
    $h += '<button data-status="diff">Unterschiedlich (' + $Diff + ')</button>'
    $h += '<button data-status="same">Identisch (' + $Same + ')</button>'
    $h += '</div><span class="info count"></span></div>'
    return $h
}

function New-SizeCell {
    param($Bytes, [string]$Text)
    if ($null -eq $Bytes) { return '<td class="num missing" data-sort="-1">fehlt</td>' }
    return '<td class="num" data-sort="' + $Bytes + '">' + $Text + '</td>'
}

function New-MarkCell {
    param([int]$Count)
    if ($Count -le 0) { return '<td class="num" data-sort="0">&ndash;</td>' }
    return '<td class="num missing" data-sort="' + $Count + '">' + $Count + '</td>'
}

function New-DiffCell {
    param($Bytes)
    if ($null -eq $Bytes) { return '<td class="num" data-sort="0">&ndash;</td>' }
    $cls = ''
    $sign = ''
    if ($Bytes -lt 0) { $cls = 'neg'; $sign = '-' }
    elseif ($Bytes -gt 0) { $cls = 'pos'; $sign = '+' }
    return '<td class="num ' + $cls + '" data-sort="' + $Bytes + '">' + $sign + (Format-Size ([math]::Abs($Bytes))) + '</td>'
}

$sb = New-Object System.Text.StringBuilder

$meta = ' &middot; Filter: <code>' + (ConvertTo-HtmlText $Filter) + '</code>'
if ($ExcludeFolder.Count -gt 0) { $meta += ' &middot; Ausgeschlossen: <code>' + (ConvertTo-HtmlText ($ExcludeFolder -join ', ')) + '</code>' }
if ($CompareHash) { $meta += ' &middot; Hash-Vergleich: ' + $HashAlgorithm }
$meta += ' &middot; Datumstoleranz: ' + $TimeToleranceSeconds + ' s'

$null = $sb.AppendLine('<!DOCTYPE html>')
$null = $sb.AppendLine('<html lang="de"><head><meta charset="utf-8">')
$null = $sb.AppendLine('<meta name="viewport" content="width=device-width, initial-scale=1">')
$null = $sb.AppendLine('<title>Ordnervergleich</title>')
$null = $sb.AppendLine('<style>' + $css + '</style>')
$null = $sb.AppendLine('</head><body>')
$null = $sb.AppendLine('<h1>Ordnervergleich</h1>')
$null = $sb.AppendLine('<div class="sub">Erstellt am ' + (Get-Date -Format 'dd.MM.yyyy HH:mm:ss') + $meta + '</div>')

$null = $sb.AppendLine('<div class="paths">')
$null = $sb.AppendLine('<div class="pathbox"><span class="tag tagA">A</span><code>' + (ConvertTo-HtmlText $idxA.Root) +
    '</code><div class="sub" style="margin:6px 0 0">' + $idxA.Files.Count + ' Dateien &middot; ' + (Format-Size $totalA) + '</div></div>')
$null = $sb.AppendLine('<div class="pathbox"><span class="tag tagB">B</span><code>' + (ConvertTo-HtmlText $idxB.Root) +
    '</code><div class="sub" style="margin:6px 0 0">' + $idxB.Files.Count + ' Dateien &middot; ' + (Format-Size $totalB) + '</div></div>')
$null = $sb.AppendLine('</div>')

# Warnung, wenn eine Seite unvollstaendig eingelesen wurde. Das ist die haeufigste
# Ursache fuer scheinbar falsche Ergebnisse: was auf einer Seite nicht gelesen
# werden konnte, erscheint zwangslaeufig als "nur auf der anderen Seite vorhanden".
$errListA = @($idxA.Errors | Where-Object { $_ -like '`[A`]*' })
$errListB = @($idxB.Errors | Where-Object { $_ -like '`[B`]*' })

if ($errListA.Count -gt 0 -or $errListB.Count -gt 0) {
    $null = $sb.AppendLine('<div class="warn">')
    $null = $sb.AppendLine('<h3>Achtung: eine Seite konnte nicht vollst&auml;ndig gelesen werden</h3>')
    $null = $sb.AppendLine('<p>Dateien, die beim Einlesen &uuml;bersprungen wurden, fehlen im Vergleich. ' +
        'Sie erscheinen dann f&auml;lschlich als <strong>nur auf der jeweils anderen Seite vorhanden</strong>, ' +
        'obwohl es sie auf beiden gibt.</p><ul>')
    if ($errListA.Count -gt 0) {
        $null = $sb.AppendLine('<li>Seite <strong>A</strong>: ' + $errListA.Count +
            ' &Uuml;bersprungene &mdash; betroffene Dateien k&ouml;nnen f&auml;lschlich als &bdquo;Nur in B&ldquo; erscheinen</li>')
    }
    if ($errListB.Count -gt 0) {
        $null = $sb.AppendLine('<li>Seite <strong>B</strong>: ' + $errListB.Count +
            ' &Uuml;bersprungene &mdash; betroffene Dateien k&ouml;nnen f&auml;lschlich als &bdquo;Nur in A&ldquo; erscheinen</li>')
    }
    $null = $sb.AppendLine('</ul><p style="margin-top:8px">H&auml;ufigste Gr&uuml;nde: Pfade &uuml;ber 260 Zeichen ' +
        '(betrifft nur die tiefer verschachtelte Seite), fehlende Zugriffsrechte, oder verkn&uuml;pfte Ordner ' +
        '(Junctions, Symlinks), denen PowerShell 5.1 nicht folgt. Die Einzelheiten stehen unten unter ' +
        '&bdquo;Hinweise&ldquo;.</p>')
    $null = $sb.AppendLine('</div>')
}

$null = $sb.AppendLine('<div class="cards">')
$null = $sb.AppendLine('<div class="card onlyA"><div class="k">Dateien nur in A</div><div class="v">' + $fOnlyA + '</div></div>')
$null = $sb.AppendLine('<div class="card onlyB"><div class="k">Dateien nur in B</div><div class="v">' + $fOnlyB + '</div></div>')
$null = $sb.AppendLine('<div class="card diff"><div class="k">Unterschiedlich</div><div class="v">' + $fDiff + '</div></div>')
$null = $sb.AppendLine('<div class="card same"><div class="k">Identisch</div><div class="v">' + $fSame + '</div></div>')
$null = $sb.AppendLine('<div class="card"><div class="k">Ordner gepr&uuml;ft</div><div class="v">' + $dirRows.Count + '</div></div>')
$null = $sb.AppendLine('<div class="card"><div class="k">Dateien gepr&uuml;ft</div><div class="v">' + $fileRows.Count + '</div></div>')
$null = $sb.AppendLine('</div>')

# ---- Tabelle 1: Verzeichnisse ----
$null = $sb.AppendLine('<h2>1. Verzeichnisse</h2>')
$null = $sb.AppendLine('<div class="panel">')
$null = $sb.AppendLine((New-Toolbar -OnlyA $dOnlyA -OnlyB $dOnlyB -Diff $dDiff -Same $dSame))
$null = $sb.AppendLine('<div class="scroll"><table><thead><tr>')
$null = $sb.AppendLine('<th>Status</th><th>Verzeichnis (relativ)</th><th class="num">Gr&ouml;&szlig;e A</th><th class="num">Gr&ouml;&szlig;e B</th><th class="num">Differenz</th><th class="num">Dateien A</th><th class="num">Dateien B</th><th class="num" title="Dateien darunter, die in B fehlen">Fehlt in B</th><th class="num" title="Dateien darunter, die in A fehlen">Fehlt in A</th><th class="num" title="Dateien darunter, die auf beiden Seiten liegen, aber abweichen">Abweichend</th>')
$null = $sb.AppendLine('</tr></thead><tbody>')

foreach ($d in $dirRows) {
    $k = Get-StatusKey $d.Status
    $search = ($d.RelativerPfad + ' ' + $d.Name).ToLower()

    $cA = '<td class="num missing" data-sort="-1">&ndash;</td>'
    if ($null -ne $d.DateienA) { $cA = '<td class="num" data-sort="' + $d.DateienA + '">' + $d.DateienA + '</td>' }
    $cB = '<td class="num missing" data-sort="-1">&ndash;</td>'
    if ($null -ne $d.DateienB) { $cB = '<td class="num" data-sort="' + $d.DateienB + '">' + $d.DateienB + '</td>' }

    $anzeigePfad = $d.RelativerPfad
    if ($anzeigePfad -eq '.') { $anzeigePfad = '(Wurzel)' }

    $null = $sb.AppendLine('<tr data-status="' + $k + '" data-search="' + (ConvertTo-HtmlText $search) + '">' +
        '<td data-sort="' + $d.Status + '"><span class="badge b-' + $k + '">' + (ConvertTo-HtmlText $d.Status) + '</span></td>' +
        '<td class="path">' + (ConvertTo-HtmlText $anzeigePfad) + '</td>' +
        (New-SizeCell -Bytes $d.GroesseA_Bytes -Text $d.GroesseA) +
        (New-SizeCell -Bytes $d.GroesseB_Bytes -Text $d.GroesseB) +
        (New-DiffCell -Bytes $d.Differenz_Bytes) +
        $cA + $cB +
        (New-MarkCell -Count $d.FehltInB) +
        (New-MarkCell -Count $d.FehltInA) +
        (New-MarkCell -Count $d.Abweichend) + '</tr>')
}
$null = $sb.AppendLine('</tbody></table></div></div>')

# ---- Tabelle 2: Dateien ----
$null = $sb.AppendLine('<h2>2. Dateien</h2>')
$null = $sb.AppendLine('<div class="panel">')
$null = $sb.AppendLine((New-Toolbar -OnlyA $fOnlyA -OnlyB $fOnlyB -Diff $fDiff -Same $fSame))
$null = $sb.AppendLine('<div class="scroll"><table><thead><tr>')
$null = $sb.AppendLine('<th>Status</th><th>Dateiname</th><th>Verzeichnis (relativ)</th><th class="num">Gr&ouml;&szlig;e A</th><th class="num">Gr&ouml;&szlig;e B</th><th>Ge&auml;ndert A</th><th>Ge&auml;ndert B</th><th>Unterschied</th>')
$null = $sb.AppendLine('</tr></thead><tbody>')

foreach ($f in $fileRows) {
    $k = Get-StatusKey $f.Status
    $search = ($f.RelativerPfad + ' ' + $f.Dateiname + ' ' + $f.DateinameB).ToLower()

    $dA = '<td class="dt missing">fehlt</td>'
    if ($f.GeaendertA -ne '') { $dA = '<td class="dt">' + $f.GeaendertA + '</td>' }
    $dB = '<td class="dt missing">fehlt</td>'
    if ($f.GeaendertB -ne '') { $dB = '<td class="dt">' + $f.GeaendertB + '</td>' }

    $verz = $f.Verzeichnis
    if ($verz -eq '.') { $verz = '(Wurzel)' }

    # Bei abweichender Gross-/Kleinschreibung beide Schreibweisen zeigen
    $nameZelle = ConvertTo-HtmlText $f.Dateiname
    if ($f.DateinameA -ne '' -and $f.DateinameB -ne '' -and
        -not [string]::Equals($f.DateinameA, $f.DateinameB, [StringComparison]::Ordinal)) {
        $nameZelle = 'A: ' + (ConvertTo-HtmlText $f.DateinameA) +
                     '<br>B: ' + (ConvertTo-HtmlText $f.DateinameB)
    }

    $null = $sb.AppendLine('<tr data-status="' + $k + '" data-search="' + (ConvertTo-HtmlText $search) + '">' +
        '<td data-sort="' + $f.Status + '"><span class="badge b-' + $k + '">' + (ConvertTo-HtmlText $f.Status) + '</span></td>' +
        '<td class="path">' + $nameZelle + '</td>' +
        '<td class="path">' + (ConvertTo-HtmlText $verz) + '</td>' +
        (New-SizeCell -Bytes $f.GroesseA_Bytes -Text $f.GroesseA) +
        (New-SizeCell -Bytes $f.GroesseB_Bytes -Text $f.GroesseB) +
        $dA + $dB +
        '<td>' + (ConvertTo-HtmlText $f.Unterschied) + '</td></tr>')
}
$null = $sb.AppendLine('</tbody></table></div></div>')

# ---- Lesefehler ----
$allErrors = @($idxA.Errors) + @($idxB.Errors)
if ($allErrors.Count -gt 0) {
    $null = $sb.AppendLine('<h2>Hinweise / Lesefehler (' + $allErrors.Count + ')</h2>')
    # Eigene Klasse: der Hinweis-Kasten ist keine filterbare Tabelle
    $null = $sb.AppendLine('<div class="notes" style="padding:12px">')
    foreach ($e in ($allErrors | Select-Object -First 100)) {
        $null = $sb.AppendLine('<div class="path">' + (ConvertTo-HtmlText $e) + '</div>')
    }
    $null = $sb.AppendLine('</div>')
}

$dauer = New-TimeSpan -Start $startTime -End (Get-Date)
$null = $sb.AppendLine('<footer>Laufzeit: ' + ('{0:hh\:mm\:ss}' -f $dauer) + ' &middot; CSV-Dateien: ' +
    (ConvertTo-HtmlText (Split-Path $csvAll -Leaf)) + ', ' +
    (ConvertTo-HtmlText (Split-Path $csvDirs -Leaf)) + ', ' +
    (ConvertTo-HtmlText (Split-Path $csvFiles -Leaf)) + '</footer>')
$null = $sb.AppendLine('<script>' + $js + '</script>')
$null = $sb.AppendLine('</body></html>')

$htmlPath = Join-Path $OutputFolder "Vergleich_$stamp.html"
[System.IO.File]::WriteAllText($htmlPath, $sb.ToString(), (New-Object System.Text.UTF8Encoding($true)))

# ----------------------------------------------------------------------------
# Zusammenfassung
# ----------------------------------------------------------------------------

Write-Host ''
Write-Host '--- Ergebnis ---------------------------------------------------' -ForegroundColor White
Write-Host ('Verzeichnisse : {0} gesamt | nur A: {1} | nur B: {2} | abweichend: {3} | identisch: {4}' -f $dirRows.Count, $dOnlyA, $dOnlyB, $dDiff, $dSame)
Write-Host ('Dateien       : {0} gesamt | nur A: {1} | nur B: {2} | abweichend: {3} | identisch: {4}' -f $fileRows.Count, $fOnlyA, $fOnlyB, $fDiff, $fSame)

if ($errListA.Count -gt 0 -or $errListB.Count -gt 0) {
    Write-Host ''
    Write-Host 'ACHTUNG: eine Seite konnte nicht vollstaendig gelesen werden.' -ForegroundColor Red
    if ($errListA.Count -gt 0) {
        Write-Host ("  Seite A: {0} uebersprungen -> koennen faelschlich als 'Nur in B' erscheinen" -f $errListA.Count) -ForegroundColor Yellow
    }
    if ($errListB.Count -gt 0) {
        Write-Host ("  Seite B: {0} uebersprungen -> koennen faelschlich als 'Nur in A' erscheinen" -f $errListB.Count) -ForegroundColor Yellow
    }
    Write-Host '  Einzelheiten stehen im HTML-Bericht unter "Hinweise".' -ForegroundColor DarkGray
}
Write-Host ''
Write-Host ('HTML : {0}' -f $htmlPath) -ForegroundColor Green
Write-Host ('CSV  : {0}' -f $csvAll)   -ForegroundColor Green
Write-Host ('CSV  : {0}' -f $csvDirs)  -ForegroundColor Green
Write-Host ('CSV  : {0}' -f $csvFiles) -ForegroundColor Green

if ($OpenReport) { Start-Process $htmlPath }

[pscustomobject]@{
    HtmlReport       = $htmlPath
    CsvGesamt        = $csvAll
    CsvVerzeichnisse = $csvDirs
    CsvDateien       = $csvFiles
    AnzahlOrdner     = $dirRows.Count
    AnzahlDateien    = $fileRows.Count
}
