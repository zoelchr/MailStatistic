<#
.SYNOPSIS
    Erstellt eine statistische Auswertung von E-Mails in einem Outlook-Postfach und exportiert die Ergebnisse in eine Excel-Datei mit Makro-Funktionalität.

.DESCRIPTION
    Dieses PowerShell-Skript durchsucht rekursiv ein ausgewähltes Outlook-Postfach und erfasst E-Mails ab einem bestimmten Zeitpunkt.
    Es extrahiert dabei Absender, Betreff, Ordnerpfad, Empfänger und weitere Metadaten jeder E-Mail.
    Die Ausgabe erfolgt in einer Excel-Datei (basierend auf einer Makro-fähigen Vorlage), in der über ein eingebettetes Makro jede E-Mail per Klick geöffnet werden kann.

    Zu den Hauptfunktionen zählen:
    - Auswahl des Outlook-Postfachs über Parameter oder interaktiv
    - Fortschrittsanzeige beim Scannen der Ordner (abschaltbar)
    - Testmodus zur Begrenzung auf eine bestimmte Anzahl E-Mails
    - Ausgabe in Excel mit optisch klickbarem "Open"-Link
    - Anpassung von Spaltenbreiten, Sortierung nach Datum, Ausblenden technischer Spalten
    - Empfängeranalyse für "To"-Feld

.PARAMETER Mailbox
    Name des Outlook-Postfachs (optional; ansonsten interaktive Auswahl).

.PARAMETER Template
    Pfad zur Excel-Vorlagendatei (.xlsm), die als Template dient.

.PARAMETER OutDir
    Verzeichnis, in dem die Ausgabedatei gespeichert wird.

.PARAMETER YearsBack
    Anzahl der Jahre, die vom heutigen Datum zurückgerechnet werden, um den Scan-Zeitraum festzulegen.

.PARAMETER MonthsBack
    Anzahl der zusätzlichen Monate für die Rückrechnung.

.PARAMETER NoProgress
    Deaktiviert die Fortschrittsanzeige.

.PARAMETER Testing
    Aktiviert den Testmodus mit reduzierter Anzahl verarbeiteter Mails (Standard: 100).

.NOTES
    Autor: Rüdiger Zölch
    Lizenz: MIT License (siehe unten)
    Erstellt: 2025
    Kompatibilität: Windows PowerShell 5.1 und Microsoft Outlook (COM-Schnittstelle)

.LICENSE
    MIT License

    Copyright (c) 2025 Rüdiger Zölch

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE.
#>

param(
    [string]$Mailbox    = "Shapth-Projekt (LGL)",
    [string]$Template,
    [string]$OutDir,
    [int]   $YearsBack  = 0,
    [int]   $MonthsBack = 1, 		# Standardmäßig werden die Emails seit dem letzten Monat abgefragt
    [string]$StartDate,
    [string]$EndDate,
	[switch]$NoProgress,            # Fortschrittsanzeige deaktivierbar
	[switch]$Testing,         		# Testmodus mit Begrenzung maximale Emailanzahl für schnelleren Durchlauf
    [switch]$NoMailboxquery         # Ohne User-Abfrage der zu verwendenden Mailbox 
)

# Nach dem param-Block: $ScriptDir berechnen
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Standardpfade nur setzen, wenn leer übergeben
if (-not $Template) {
    $Template = Join-Path $ScriptDir "MailStatisticTemplate.xlsm"
}
if (-not $OutDir) {
    $OutDir = Join-Path $ScriptDir "Export"
}

# Export-Verzeichnis erzeugen, wenn noch nicht vorhanden
if (-not (Test-Path $OutDir)) {
    New-Item -Path $OutDir -ItemType Directory | Out-Null
}

# Ausgabe zur Kontrolle
Write-Host "Script-Verzeichnis: $ScriptDir"
Write-Host "Pfad zur Vorlage:   $Template"
Write-Host "Exportverzeichnis:  $OutDir"

# Verkürzter Lauf im Testfall
$TestEmailCount = 20 # Anzahl der maximal exportierten Emails
if ($Testing) {
    Write-Host "Testmodus aktiv: Es werden nur maximal $TestEmailCount Mails verarbeitet." -ForegroundColor Yellow
}

# Outlook-Objektmodell initialisieren

# Start (bzw. Verbindung mit) Outlook über die Windows-COM-Schnittstelle.
$ol = New-Object -ComObject Outlook.Application

# MAPI-Namespace holne, also die Struktur der E-Mail-Ordner, Kalender, Kontakte usw. in Outlook.
# $ns ist jetzt das Einstiegspunkt-Objekt in die Postfachhierarchie.
$ns = $ol.GetNamespace("MAPI")

# ---------- 0  Mailbox abfragen ---------------------------------------
if (-not $NoMailboxquery) {
    # Alle verfügbaren Postfächer ermitteln
    $mailboxes = @()
    for ($i = 1; $i -le $ns.Folders.Count; $i++) {
        $mailboxes += $ns.Folders.Item($i).Name
    }

    # Liste anzeigen
    Write-Host "`nVerfügbare Postfächer:`n"
    for ($i = 0; $i -lt $mailboxes.Count; $i++) {
        Write-Host "$($i + 1): $($mailboxes[$i])"
    }

    # Auswahl abfragen
    do {
        $selection = Read-Host "Bitte gib die Nummer des gewünschten Postfachs ein"
    } while (-not ($selection -match '^\d+$') -or [int]$selection -lt 1 -or [int]$selection -gt $mailboxes.Count)

    # Ausgewähltes Postfach speichern
    $Mailbox = $mailboxes[[int]$selection - 1]

    Write-Host "`nAusgewähltes Postfach: $Mailbox"
}

# ---------- 1  Vorlage prüfen -----------------------------------------
if (-not (Test-Path $Template)) {
    Write-Error "Vorlagendatei nicht gefunden: $Template"
    exit
}

# ---------- 2  Datumspanne bestimmen ----------------------------------

# Zuerst wird geprüft, ob das Attribut StartDate im Format yyyy-mm-dd übergeben wurde.
try {
    if ($StartDate) {
        $StartDate = [datetime]::ParseExact($StartDate, 'yyyy-MM-dd', $null)
    } else {
        $StartDate = [DateTime]::MinValue # Wenn kein Parsen möglich, dann wird StartDate auf das frühest mögliche Datum gesetzt
    }

    # Dann wird geprüft, ob das Attribut EndDate im Format yyyy-mm-dd übergeben wurde.
    if ($EndDate) {
        $EndDate = [datetime]::ParseExact($EndDate, 'yyyy-MM-dd', $null)
    } else {
        $EndDate = Get-Date # Wenn kein Parsen möglich, dann wird EndDate auf das aktuelle Datum gesetzt
    }
} catch {
    Write-Host "FEHLER: Ungültiges Datumsformat. Bitte verwende 'yyyy-MM-dd'." -ForegroundColor Red
    exit
}

# Wenn noch kein StartDate definiert wurde, dann erfolgt Auswertung der Attribute YearsBack und MonthsBack
if (-not $StartDate) {

    $StartDate = if ($YearsBack -eq 0 -and $MonthsBack -eq 0) {
                    [DateTime]::MinValue				 
                } else {
                    (Get-Date).AddYears(-$YearsBack).AddMonths(-$MonthsBack)
                }
}

Write-Host "Ich exprotiere die Emails ab: $StartDate bis $EndDate" # Ausgabe zur Kontrolle

# ---------- 3  Mails einsammeln ---------------------------------------
Remove-Variable -Name stats,seen -ErrorAction SilentlyContinue
# Erzeugung eine dynamisch wachsende ArrayList zur Speicherung von E-Mail-Statistikdaten (jeweils als Objekt)
$stats = [System.Collections.ArrayList]::new() 

# Erzeugung eines HashSet für Strings zur Duplikatserkennung
$seen  = New-Object 'System.Collections.Generic.HashSet[string]'

# Hilfsfunktion, die ein neues E-Mail-Statistikdaten-Objekt zur Liste $stats hinzufügt.
# Die Zuweisung an $null unterdrückt die Ausgabe im Terminal.
function Add-Stat([object]$o){ $null = $stats.Add($o) }

# Suche im obersten Ordner-Level nach einem Postfach mit dem Namen $Mailbox.
# Falls kein solches Postfach gefunden wird ($mbx ist $null), wird mit throw eine Fehlermeldung erzeugt und das Skript bricht kontrolliert ab.
$mbx = $ns.Folders.Item($Mailbox)  ; if(!$mbx){throw "Mailbox not found."}

# Liste der nicht relevanten Ordner
$skipFolders = @("Kontakte", "Kalender", "Aufgaben", "Journal", "Notes", "Deleted Items", "Gelöschte Objekte", "Yammer", "Kalender & Abwesenheiten", "Entwürfe") 

$global:TestCounter = 0
$global:FolderCounter = 0
$global:ItemCounter = 0

<#
.SYNOPSIS
    Durchsucht rekursiv alle E-Mail-Ordner eines Outlook-Postfachs und sammelt Statistikdaten zu gesendeten E-Mails.

.DESCRIPTION
    Die Funktion `Scan` durchsucht einen angegebenen Outlook-Ordner (und dessen Unterordner), filtert dabei ausschließlich gültige E-Mail-Elemente 
    vom Typ 'IPM.Note*' und berücksichtigt nur E-Mails, deren Versanddatum (`SentOn`) nach dem global definierten Stichtag `$StartDate` liegt.

    Duplikate werden vermieden, indem jede E-Mail anhand ihrer eindeutigen EntryID in einem HashSet `$seen` überprüft wird. Für jede gültige und neue 
    E-Mail wird ein PowerShell-CustomObject mit relevanten Metadaten erstellt und der globalen Statistikliste `$stats` hinzugefügt.

    Optional kann die Verarbeitung im Testmodus (`$Testing`) nach einer vordefinierten Anzahl von E-Mails (`$TestEmailCount`) vorzeitig beendet werden.

    Zusätzlich werden:
    - Fortschrittsanzeigen (via `Write-Progress`) sowohl auf Ordnerebene als auch auf Elementebene angezeigt, sofern nicht mit `-NoProgress` deaktiviert.
    - Nicht-relevante Ordner (z.B. Kalender, Kontakte) sowie individuell ausgeschlossene Ordner über `$skipFolders` ignoriert.
    - Empfängeradressen im Feld `Recipients` gespeichert, jedoch nur solche vom Typ "To".

.PARAMETER $fld
    Outlook-Folder-Objekt, das als Einstiegspunkt für die rekursive Verarbeitung dient.

.OUTPUTS
    Gibt `$true` zurück, wenn die Verarbeitung im Testmodus vorzeitig beendet wurde. Ansonsten `$false`.

.NOTES
    Die gesammelten Objekte werden nicht direkt zurückgegeben, sondern der Liste `$stats` hinzugefügt.
    Die Funktion verwendet globale Variablen: `$stats`, `$seen`, `$StartDate`, `$Testing`, `$TestCounter`, `$TestEmailCount`, `$FolderCounter`, `$NoProgress`, `$skipFolders`.

#>
function Scan($fld) {
    $olMailItem = 0 # Nur IPM.Note*-Objekte, also echte E-Mails.
	
	# Überspringe Ordner, die keine E-Mails enthalten (z.B. Kalender, Kontakte)
    if ($fld.DefaultItemType -ne $olMailItem) {
        return $false
    }
	
	# Überspringe Ordner, die in der Liste der nicht relevanten Ordner vorkommen
	if ($skipFolders -contains $fld.Name) {
		return $false
	}

	# Erhöhe globalen Zähler für besuchte Ordner (für Fortschrittsanzeige)
    $global:FolderCounter++
	
	# Zeige Fortschritt, wenn nicht per -NoProgress unterdrückt
    if (-not $NoProgress) {
        Write-Progress -Activity "Outlook-Scan läuft..." `
                       -Status "Ordner: $($fld.FolderPath)" `
                       -PercentComplete ($global:FolderCounter % 100)
    }

	# Durchlaufe alle Elemente im aktuellen Ordner
	$total = $fld.Items.Count # Anzahl der Elemente im aktuellen Ordner
	$i = 0 # Initialisiere Counter für Fortschrittsanzeige

    foreach ($itm in $fld.Items) {
		
		$i++
		if (-not $NoProgress) {
			Write-Progress 	-Activity "Verarbeite Ordner..." `
							-Status "$($fld.FolderPath): $i von $total" `
							-PercentComplete ([math]::Round(($i / $total) * 100))
		}
		
		# Verarbeite nur echte E-Mails (nicht z.B. Termine, Besprechungsanfragen)
        if ($itm.MessageClass -like 'IPM.Note*' -and $itm.SentOn -ge $StartDate) {
			
			# Vermeide Duplikate: prüfe, ob EntryID schon bekannt
			# Füge die EntryID der E-Mail zum Set $seen hinzu – und fahre nur fort, wenn sie dort noch nicht enthalten war.
			# $seen ist eine sogenannte HashSet-Datenstruktur, doppelte Werte werden automatisch ignoriert.
			# $seen.Add(...) gibt $true zurück, wenn der Eintrag neu ist (also noch nicht im Set vorhanden war).
            if ($seen.Add($itm.EntryID)) {	

				# Neue Statistikzeile hinzufügen
				Add-Stat([pscustomobject]@{
					SentOn    = $itm.SentOn															# Datum/Zeit
					Sender    = $itm.SenderName 													# Absender
					BehalfOf  = $itm.SentOnBehalfOfName												# Gesendet im Auftrag von (optional)
					Subject   = if ($itm.Subject) { $itm.Subject } else { '(no subject)' }			# Betreff
					Folder    = $fld.FolderPath														# Ordnerpfad
					Words     = if ($itm.Body) { ($itm.Body -split '\s+').Count } else { 0 }		# Wortanzahl im Body
					StoreID   = $fld.StoreID														# Eindeutige ID des Postfachs
					EntryID   = $itm.EntryID														# Eindeutige ID der Email 
					OpenTxt   = 'Open'																# Wird später als Hyperlink mit Hilfe eines VBA-Makroks verwendet
					Recipients = (
						$itm.Recipients | 
						Where-Object { $_.Type -eq 1 } | 											# Nur "To"-Empfänger
						ForEach-Object { $_.Name }) -join "; "										# Empfängername(n) als String
				})

				# Testmodus: Brich nach X Mails ab
                if ($Testing) {
                    $global:TestCounter++
                    if ($global:TestCounter -ge $TestEmailCount) {
                        return $true
                    }
                }
            }
        }
    }

	# Rekursiver Aufruf für alle Unterordner
    foreach ($s in $fld.Folders) {
        if (Scan $s) {
            return $true
        }
    }

    return $false # Normales Ende
}

# Aufruf der Scan-Funktion mit dem Postfach mit dem Namen $Mailbox
Scan $mbx | Out-Null

# Abbruch, wenn keine Emails gefunden wurden
if($stats.Count -eq 0){Write-Warning 'No mails found.';return}

# ---------- 4  Vorlage kopieren & Daten schreiben ----------------------

# Namen der Excel-Export-Datei erzeugen
$timestamp = Get-Date -f 'yyyyMMdd_HHmmss'
$outName = "MailStatistic_$timestamp.xlsm"
$outFile   = Join-Path $OutDir "MailStatistic_$timestamp.xlsm"

# Start (bzw. Verbindung mit) Excel über die Windows-COM-Schnittstelle.
$xl = New-Object -ComObject Excel.Application

# Excel im Hintergrund laufen lassen (nicht sichtbar für den Benutzer).
# Das beschleunigt das Skript und vermeidet visuelle Störungen.
$xl.Visible = $false

# Öffnet die Excel-Datei, die als Vorlage dient (z. B. mit vordefiniertem Makro und Formatierung).
# Parameter:
# 1. $Template – Pfad zur Vorlagendatei
# 2. $null – keine speziellen Update-Einstellungen
# 3. $true – Datei wird im **Nur-Lesen-Modus** geöffnet (zum Schutz der Vorlage)
$wb = $xl.Workbooks.Open($Template,$null,$true)

# Speichert eine **Kopie** der geöffneten Vorlage unter dem Namen `$outFile`.
# Das bedeutet: Die Vorlage selbst bleibt unverändert, es wird nur mit einer Kopie gearbeitet.
$wb.SaveCopyAs($outFile)

# Schließt die geöffnete Vorlage (die im Nur-Lesen-Modus geöffnet war), ohne sie zu speichern.
$wb.Close($false)

# Öffnet jetzt die gespeicherte Kopie als **neue Arbeitsmappe**, diesmal **nicht** im Nur-Lesen-Modus.
# In dieser Datei werden im weiteren Verlauf Daten eingetragen.
$wb2 = $xl.Workbooks.Open($outFile)

# Referenziert das erste Arbeitsblatt in der geöffneten Arbeitsmappe.
$ws   = $wb2.Worksheets.Item(1)


# Postfachname als Blattname verwenden (max. 31 Zeichen, keine Sonderzeichen)
$sheetName = $Mailbox -replace '[:\\/*?\[\]]', ''  # Entfernt unerlaubte Zeichen
$sheetName = $sheetName.Substring(0, [Math]::Min(31, $sheetName.Length))  # Kürzen auf 31 Zeichen

try {
    $ws.Name = $sheetName
} catch {
    Write-Warning "Blattname '$sheetName' konnte nicht gesetzt werden (möglicherweise schon vergeben)."
}

# Alle E-Mails aus der ArrayList abarbeiten

$row = 2 # Die erste Zeile sind die Spaltenüberschriften

foreach ($entry in $stats) {
	$ws.Cells($row,1).Value = $entry.StoreID
    $ws.Cells($row,2).Value = $entry.EntryID

    $cell = $ws.Cells($row, 3)
    $cell.Value = "Open"
    $cell.Font.Color = 16711680   # Blau
    $cell.Font.Underline = 2

    $ws.Cells($row,4).Value = $entry.SentOn.ToString('yyyy-MM-dd HH:mm')
    $ws.Cells($row,5).Value = $entry.Sender
    $ws.Cells($row,6).Value = $entry.BehalfOf
    $ws.Cells($row,7).Value = $entry.Subject
    $ws.Cells($row,8).Value = $entry.Folder
    $ws.Cells($row,9).Value = "$($entry.Words)"
    $ws.Cells($row,10).Value = $entry.Recipients

	# Einen zusätzlichen Vergleichsschlüssel erzeugen, um unerkannte Dobletten zu eleminieren
	$comparisonKey = (
		$entry.Subject + '|' +
		$entry.SenderName + '|' +
		$entry.SentOn.ToString("yyyy-MM-dd HH:mm")		
	).ToLower()
	#$comparisonKey = "$($entry.SentOn.Ticks)-$($entry.Sender)-$($entry.Subject)"
	$ws.Cells($row, 11).Value2 = $comparisonKey
	
    $row++
}

# Doublettenerkennung: Alle mehrfach vorkommenden Vergleichsschlüssel finden
$comparisonKeys = @{}
$rowCount = $ws.UsedRange.Rows.Count

for ($r = 2; $r -le $rowCount; $r++) {
    $key = $ws.Cells($r, 11).Text
    if ([string]::IsNullOrWhiteSpace($key)) { continue }

    if ($comparisonKeys.ContainsKey($key)) {
        # Schon mal gesehen → markiere als Doublette in Spalte 12
        $ws.Cells($r, 12).Value2 = "Doublette"
    } else {
        # Erster Fund → nur merken
        $comparisonKeys[$key] = $true
    }
}

# Autofilter setzen auf Zeile 1, Bereich von Spalte A bis Spalte L (12 Spalten)
$ws.Range("A1:L1").AutoFilter() | Out-Null

# Filter für Spalte 12 (Index 12): Nur Zellen anzeigen, die NICHT "Doublette" enthalten
# (Kriterium "<>Doublette" bedeutet: alles außer "Doublette")
$ws.Range("A1:L$rowCount").AutoFilter(12, "<>Doublette") | Out-Null

# Grafische Formatierung inklusive Datenfilter je Spalte 
$ws.ListObjects.Add(1,$ws.Range("A1").CurrentRegion,$null,1)|Out-Null

# Automatische Einstellung der Spaltenbreite
$ws.UsedRange.Columns.AutoFit()|Out-Null

# Die beiden ersten Spalten (StoreID und EntryID) ausblenden
$ws.Columns("A:A").Hidden = $true
$ws.Columns("B:B").Hidden = $true
$ws.Columns("K:K").Hidden = $true

# Spaltenbreite vorgeben
$ws.Columns.Item(3).ColumnWidth = 10   # Open Email
$ws.Columns.Item(5).ColumnWidth = 30   # Sender
$ws.Columns.Item(6).ColumnWidth = 30   # BehalfOf
$ws.Columns.Item(7).ColumnWidth = 70   # Subject
$ws.Columns.Item(8).ColumnWidth = 60   # Folder
$ws.Columns.Item(9).ColumnWidth = 9   	# Words
$ws.Columns.Item(10).ColumnWidth = 90  # Direct Recipients

# Bereich für Sortierung festlegen
$usedRange = $ws.UsedRange
$sortRange = $usedRange.Resize($usedRange.Rows.Count, $usedRange.Columns.Count)
$sort = $ws.Sort
$sort.SortFields.Clear()

# Sortieren nach Spalte A (Datum), absteigend (neueste oben)
$sort.SortFields.Add(
    $ws.Range("D2"),      # Start der Sortierung ab zweiter Zeile
    0,                    # xlSortOnValues
    2,                    # xlDescending
    $null,
    0                     # xlSortNormal
) | Out-Null

$sort.SetRange($sortRange)
$sort.Header = 1          # xlYes (erste Zeile = Kopfzeile)
$sort.MatchCase = $false
$sort.Orientation = 1     # xlTopToBottom
$sort.Apply() | Out-Null

# Excel-Workbook speichern, schließen und Excel beenden
$wb2.Save() | Out-Null
$wb2.Close($false) | Out-Null 
$xl.Quit() | Out-Null

# COM-Objekt manuell aus dem Speicher entladen, um Ressourcen freizugeben.
[Runtime.InteropServices.Marshal]::ReleaseComObject($ol)|Out-Null

Write-Host "Done -> $outFile  ($($stats.Count) rows)"
Write-Host "Hinweis: Damit der Link auf die Emails funktioniert muss beim ersten Öffnen von $outName 'Inhalt aktivieren' bestätigt werden."
