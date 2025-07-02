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
# ────────────────────────────────────────────────────────────────────────────────────────────
# Param-Block muss zwingend am Anfang stehen
# Hinweis: 
# Diese Parameter sind auch in allen Funktion lesend sichtbar. 
# Ein Schreibzugriff auf diese Paraemter würde aber eine andere Definition erfordern, z.B. [string] $script:EXCELTEMPLATE
param(
    [string] $EXCELTEMPLATE,
    [string] $OUTDIR,
    [datetime] $STARTDATE,
    [datetime] $ENDDATE,
	[switch] $NOPROGRESS,               # Fortschrittsanzeige deaktivierbar
    [switch] $NOCONSOLELOGGING,         # Keine Debbugingausgabe auf der Konsole
    [switch] $FILELOGGING,              # Debug-File erzeugen
	[switch] $TESTING,         	        # Testmodus mit Begrenzung maximale Emailanzahl für schnelleren Durchlauf
    [switch] $NOMEAILBOXQUERY,          # Ohne User-Abfrage der zu verwendenden Mailbox 
    [string[]] $MAILBOXES,
    [int] $YEARSBACK,
    [int] $MONTHBACK 		            # Standardmäßig werden die Emails seit dem letzten Monat abgefragt
)

# Defaultwerte setzen
if (-not $MAILBOXES) { $MAILBOXES = @('Postfach A') } # Hier muss ein Postfach als Standard eingetragen werden
if (-not $YEARSBACK) {$YEARSBACK=0}
if (-not $MONTHBACK) {$MONTHBACK=1}
if (-not $script:NOCONSOLELOGGING) {$script:NOCONSOLELOGGING = $true} # Keine Ausgabe des Loggings auf der Konsole

# Bei den Switches ein Casting durchführen (alternativ wäre auch $FILELOGGING.IsPresent möglich) und im ganzen Script verfügbar machen
$script:FILELOGGING = [bool]$FILELOGGING
$script:NOCONSOLELOGGING = [bool]$script:NOCONSOLELOGGING

# Verkürzter Lauf im Testfall
$TestEmailCount = 40 # Anzahl der maximal exportierten Emails
if ($TESTING) {
    Write-Host "Testmodus aktiv: Es werden nur maximal $TestEmailCount Mails verarbeitet." -ForegroundColor Yellow
}
   
# ─────────────────────────────────Log-Variante 1──────────────────────────────────────────────                            
# Methode: Eigene Log-Funktion "Write-Log" mit Log-Datei inklusive Zeitstempel und Zeilennummer
$timestampdebuglog = Get-Date -Format 'yyyy-MM-dd_HH-mm'
$filenamedebuglog = "debug_$timestampdebuglog.log"
$Script:LogFile = Join-Path $PSScriptRoot $filenamedebuglog # Variabel ist nur im Skript verfügbar

# ─────────────────────────────────Log-Variante 2────────────────────────────────────────────── 
# Methode: Nutzung der Powershell-Lösung PS-Tracing
# Damit Trace nicht auf Konosle daregstellt wird ist folgender Aufruf erforderlich: powershell -File .\MailStatistic.ps1 5> debug.log
#Start-Transcript -Path "$PSScriptRoot\debug_full.log"   # alles mitschneiden 
#Set-PSDebug -Trace 0   # 0=aus, 1=befehle, 2+ mit Variablen

# ────────────────────────────────────────────────────────────────────────────────────────────
#region Functions

# Beispiel Aufruf: Write-Log "Postfach belegt > 90 %" -Level WARN
# 
# .\MailStatistic.ps1 -NoConsoleLogging          # DEBUG-Zeilen nicht sichtbar
function Write-Log {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string] $Message,
        [ValidateSet('INFO','WARN','ERROR','DEBUG')]
        [string] $Level = 'INFO'
    )

    $frame      = (Get-PSCallStack)[1]                  # Aufrufer
    $callerName = $frame.FunctionName
    if (-not $callerName) { $callerName = '<Script>' }

    $lineNumber = $frame.ScriptLineNumber
    if (-not $lineNumber) { $lineNumber = '?' }

    $stamp   = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $logLine = '{0} [{1}] [{2}:{3}] {4}' -f $stamp, $Level, $callerName, $lineNumber, $Message

    if (-not $script:NOCONSOLELOGGING) {
        switch ($Level) {
            'WARN'  { Write-Warning "$callerName($lineNumber): $Message" }
            'ERROR' { Write-Error   "$callerName($lineNumber): $Message" }
            'DEBUG' { Write-Verbose $Message }
            default { Write-Host    $logLine }
        }
    }

    # ---------- Datei ----------
    Add-Content -Path $Script:LogFile -Value $logLine
}

<#
.SYNOPSIS
    Durchsucht rekursiv alle E-Mail-Ordner eines Outlook-Postfachs und sammelt Statistikdaten zu gesendeten E-Mails.

.DESCRIPTION
    Die Funktion `Scan` durchsucht einen angegebenen Outlook-Ordner (und dessen Unterordner), filtert dabei ausschließlich gültige E-Mail-Elemente 
    vom Typ 'IPM.Note*' und berücksichtigt nur E-Mails, deren Versanddatum (`SentOn`) nach dem global definierten Stichtag `$STARTDATE` liegt.

    Duplikate werden vermieden, indem jede E-Mail anhand ihrer eindeutigen EntryID in einem HashSet `$seen` überprüft wird. Für jede gültige und neue 
    E-Mail wird ein PowerShell-CustomObject mit relevanten Metadaten erstellt und der globalen Statistikliste `$script:stats` hinzugefügt.

    Optional kann die Verarbeitung im Testmodus (`$TESTING`) nach einer vordefinierten Anzahl von E-Mails (`$TestEmailCount`) vorzeitig beendet werden.

    Zusätzlich werden:
    - Fortschrittsanzeigen (via `Write-Progress`) sowohl auf Ordnerebene als auch auf Elementebene angezeigt, sofern nicht mit `-NoProgress` deaktiviert.
    - Nicht-relevante Ordner (z.B. Kalender, Kontakte) sowie individuell ausgeschlossene Ordner über `$skipFolders` ignoriert.
    - Empfängeradressen im Feld `Recipients` gespeichert, jedoch nur solche vom Typ "To".

.PARAMETER $fld
    Outlook-Folder-Objekt, das als Einstiegspunkt für die rekursive Verarbeitung dient.

.OUTPUTS
    Gibt `$true` zurück, wenn die Verarbeitung im Testmodus vorzeitig beendet wurde. Ansonsten `$false`.

.NOTES
    Die gesammelten Objekte werden nicht direkt zurückgegeben, sondern der Liste `$script:stats` hinzugefügt.
    Die Funktion verwendet globale Variablen: `$script:stats`, `$seen`, `$STARTDATE`, `$TESTING`, `$TestCounter`, `$TestEmailCount`, `$FolderCounter`, `$NOPROGRESS`, `$skipFolders`.

#>
function Scan($CurrentMailbox, $fld) {

    Write-Log ("Bearbeite Ordner '{0}' im Postfach '$CurrentMailbox' ..." -f $fld.Name)
    
	# Überspringe Ordner, die in der Liste der nicht relevanten Ordner vorkommen
    If ($script:FILELOGGING) {Write-Log ("Prüfe Verzeichnis '{0}' ob Teil der Ignore-Liste." -f $fld.FolderPath)}
    if (Test-SkipFolder $fld.FolderPath) { 
        If ($script:FILELOGGING) {Write-Log ("Überspringe Ordner '{0}', da dieser in der Liste der nicht relevanten Ordner vorkommt." -f $fld.Name)}
        $script:FolderIgnoreCounter++
        return $false 
    }
    else {
        If ($script:FILELOGGING) {Write-Log ("Nein - Verzeichnis wird bearbeitet.")}
    }

	# Erhöhe globalen Zähler für besuchte Ordner (für Fortschrittsanzeige)
    $script:FolderCounter++
    If ($script:FILELOGGING) {Write-Log "FolderCounter = $script:FolderCounter"}
	
	# Zeige Fortschritt, wenn nicht per -NoProgress unterdrückt
    if (-not $NOPROGRESS) {
        Write-Progress -Activity "Outlook-Scan läuft..." `
                       -Status "Ordner: $($fld.FolderPath)" `
                       -PercentComplete ($script:FolderCounter % 100)
    }

	# Durchlaufe alle Elemente im aktuellen Ordner
	$total = $fld.Items.Count # Anzahl der Elemente im aktuellen Ordner
    Write-Log ("Im Ordner '{0}' befinden sich $total Elemente." -f $fld.Name)

	$i = 0 # Initialisiere Counter für Fortschrittsanzeige

    foreach ($itm in $fld.Items) {

        # Nur echte Mails (‘IPM.Note’) durchlassen
        if ($itm.MessageClass -ne 'IPM.Note') { 
            If ($script:FILELOGGING) {Write-Log ("Überspringe Element {0} da nicht von Typ 'IPM.Note'." -f $itm.EntryID)}
            $script:IPMNoteIgnoreCounter++
            continue 
        }
        if (-not $itm.SentOn) { 
            If ($script:FILELOGGING) {Write-Log ("Überspringe Element {0} da kein Versanddatum." -f $itm.EntryID)}
            $script:NoSentOnIgnoreCounter++
            continue
        } 

        # Versanddatum der Email in das Logfile schreiben
        If ($script:FILELOGGING) {Write-Log ("Versanddatum der Email ist {0}" -f $itm.SentOn)}

        # Fortschittsanzeige aktualisieren
		$i++
		if (-not $NOPROGRESS) {
			Write-Progress 	-Activity "Verarbeite Ordner..." `
							-Status "$($fld.FolderPath): $i von $total" `
							-PercentComplete ([math]::Round(($i / $total) * 100))
		}
		
        # Datumsfilter
        if ($itm.SentOn -lt $STARTDATE) {
            $script:ItemOlderStartdateCounter++
        }

		# Verarbeite nur echte E-Mails (nicht z.B. Termine, Besprechungsanfragen)
        if ($itm.MessageClass -like 'IPM.Note*' -and $itm.SentOn -ge $STARTDATE) {
			
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
					Mailbox   = $CurrentMailbox                                                     # Mailbox
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

                If ($script:FILELOGGING) {Write-Log ("Element {0} der Liste hinzugefügt." -f $itm.EntryID)}  

				# Testmodus: Brich nach X Mails ab
                if ($TESTING) {
                    $script:TestCounter++
                    if ($script:TestCounter -ge $TestEmailCount) {
                        If ($script:FILELOGGING) {Write-Log ("Abbruch nach $TestEmailCount E-Mails.")}
                        $script:TestCounter = 0
                        return $true
                    }
                }
            } 
            else {
                # Das Element befindet sich bereits in der HashSet-Datenstruktur $seen
                If ($script:FILELOGGING) {Write-Log ("Überspringe Element {0} da Doublette." -f $itm.EntryID)}
            }
        }
        else {
            If ($script:FILELOGGING) {
                Write-Log ("Überspringe Element {0} da nicht vom Typ 'IPM.Note*' oder Zeitstempel älter als gesuchtes Startdatum." -f $itm.EntryID)
        
            }
        }
    }

    # Rekursiver Aufruf für alle Unterordner
    foreach ($sub in $fld.Folders) {
        if (Scan -CurrentMailbox $CurrentMailbox -fld $sub ) {
            return $true
        }
    }

    return $false # Normales Ende
}

# Hilfsfunktion, die ein neues E-Mail-Statistikdaten-Objekt zur Liste $script:stats hinzufügt.
# Die Zuweisung an $null unterdrückt die Ausgabe im Terminal.
function Add-Stat([object]$o)
{ 
    $null = $script:stats.Add($o) 
}

function Format-DateString {
    param(
        [object]$dt        # statt [datetime]
    )

    if (-not $dt -or $dt -eq [datetime]::MinValue) {
        return ''
    }

    return ([datetime]$dt).ToString('yyyy-MM-dd HH:mm')
}

function Get-InboxFolder {
    param($root)

    # 1. Versuche standard­mäßig „Inbox“ / „Posteingang“ (hier können weitere Namen ergänzt)
    foreach ($name in 'Inbox','Posteingang') {
        try {
            $f = $root.Folders.Item($name)
            if ($f) { 
                If ($script:FILELOGGING) {Write-Log "Ordner '$name' gefunden."}
                return $f 
            }
        } 
        catch {}
    }

    # 2. Fallback: suche nach dem Default-Mail-Class-Ordner
    foreach ($f in $root.Folders) {
        if ($f.DefaultItemType -eq 0 -and          # 0 = olMailItem
            $f.MessageClass   -match 'IPM\.Note') {
            If ($script:FILELOGGING) {Write-Log "Alternativen Ordner '$name' gefunden."}
            return $f
        }
    }

    throw "Kein Posteingangsordner in '$($root.Name)' gefunden."
}

function Connect-Outlook {
    <#
      Stellt die COM-Verbindung zu Outlook her.
      Gibt bei Erfolg das MAPI-Namespace-Objekt zurück.
      Beendet das Skript mit Exitcode 1, falls Outlook nicht erreichbar ist.
    #>

    try {
        # Outlook starten / an bestehende Instanz andocken
        $script:ol = New-Object -ComObject Outlook.Application -ErrorAction Stop

        # Einstieg in die Postfach­struktur holen
        $ns = $script:ol.GetNamespace('MAPI')        # diese Methode wirft bereits Fehler, wenn Outlook kaputt ist

        # Erfolg loggen
        If ($script:FILELOGGING) {Write-Log "Outlook-Verbindung hergestellt" -Level INFO}

        return $ns
    }
    catch {
        Write-Log "Outlook konnte nicht gestartet werden: $_" -Level ERROR
        throw  # brich komplett ab oder exit 1 – je nach gewünschtem Verhalten
    }
}

function Test-SkipFolder([string]$path) {
    #return $script:skipFolders |
    #       Where-Object { $path -like "*$_*" } |
    #       ForEach-Object { $true }           # gibt $true, sobald ≥1 Match
    
    foreach ($keyword in $script:SkipFolders) {
        if ($path -like "*$keyword*") { return $true }
    }
    return $false
}

#endregion

Write-Log "Starte Verarbeitung"

# ──────────────────────────Param-Block in das Logfile schreibe───────────────────────────────
Write-Log "Übergebene Attribute:"
Write-Log "-------------------------------"
if ($EXCELTEMPLATE) { Write-Log "EXCELTEMPLATE = $EXCELTEMPLATE" } else {Write-Log "EXCELTEMPLATE not specified."}
if ($OUTDIR) { Write-Log "OUTDIR = $OUTDIR" } else {Write-Log "OUTDIR not specified."}
if ($STARTDATE) { Write-Log "STARTDATE = $STARTDATE" } else {Write-Log "STARTDATE not specified."}
if ($ENDDATE) { Write-Log "ENDDATE = $ENDDATE" } else {Write-Log "ENDDATE not specified."}
if ($NOPROGRESS) { Write-Log "NOPROGRESS = $NOPROGRESS" } else {Write-Log "NOPROGRESS not specified."}
if ($TESTING) { Write-Log "TESTING = $TESTING" } else {Write-Log "TESTING not specified."}
Write-Log "FILELOGGING = $script:FILELOGGING" # Ich FILELOGGING gecastet
if ($script:NOCONSOLELOGGING) { Write-Log "NOCONSOLELOGGING = $script:NOCONSOLELOGGING" } else {Write-Log "NOCONSOLELOGGING not specified."}
if ($NOMEAILBOXQUERY) { Write-Log "NOMAILBOXQUERY = $NOMEAILBOXQUERY" } else {Write-Log "NOMAILBOXQUERY not specified."}
if ($MAILBOXES) { Write-Log "MAILBOXES = $MAILBOXES" } else {Write-Log "MAILBOXES not specified."}
if ($YEARSBACK) { Write-Log "YEARSBACK = $YEARSBACK" } else {Write-Log "YEARSBACK not specified."}
if ($MONTHBACK) { Write-Log "MONTHBACK = $MONTHBACK" } else {Write-Log "MONTHBACK not specified."}
Write-Log "-------------------------------"

# Falls der Aufrufer doch nur einen einzelnen, kommaseparierten String übergibt
if ($MAILBOXES.Count -eq 1 -and $MAILBOXES[0] -like '*,*') {
    $MAILBOXES = $MAILBOXES[0].Split(',') | ForEach-Object { $_.Trim() }
}

# ────────────────────────────────Counter initialisiern───────────────────────────────────────
$script:FolderIgnoreCounter = 0
$script:IPMNoteIgnoreCounter = 0
$script:NoSentOnIgnoreCounter = 0
$script:ItemOlderStartdateCounter = 0
$script:DoublettenCounter = 0

# ──────────────────────────────Relevante Pfade ermittel──────────────────────────────────────

# Pfad für die Logging-Datei
Write-Log("Vollständiger Pfad Log-File ist '$Script:LogFile'")

# $ScriptDir berechnen, also das Verzeichnis wo das Skript aufgerufen wird
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Write-Log "Skript-Verzeichnis $ScriptDir"

# Standardpfade für das Excel-Template setzen, wenn leer übergeben
if (-not $EXCELTEMPLATE) {
    $EXCELTEMPLATE = Join-Path $ScriptDir "MailStatisticTemplate.xlsm"
}

# Pfad für Ausgabe der Excel-Exporte festlegen
if (-not $OUTDIR) {
    $OUTDIR = Join-Path $ScriptDir "Export"
}

# Export-Verzeichnis erzeugen, wenn noch nicht vorhanden
if (-not (Test-Path $OUTDIR)) {
    New-Item -Path $OUTDIR -ItemType Directory | Out-Null
}

# Ausgabe zur Kontrolle
Write-Log "Script-Verzeichnis: $ScriptDir"
Write-Log "Pfad zur Vorlage:   $EXCELTEMPLATE"
Write-Log "Exportverzeichnis:  $OUTDIR"


# ──────────────────────────Outlook-Objektmodell initialisieren───────────────────────────────

# "NameSpace“ – das Outlook-Objekt
$ns = Connect-Outlook

# ──────────────────────────────────Mailboxen ermittel─────────────────────────────────────────

If ($script:FILELOGGING) {Write-Log "Ermittlung der zu durchsuchenden Mailboxen starten..."}

# Die User-Abfrage der Mailboxen erfolgt nur wenn die Variable $NOMEAILBOXQUERY nicht existiert
if (-not $NOMEAILBOXQUERY) {

    If ($script:FILELOGGING) {Write-Log "Da NOMEAILBOXQUERY = $NOMEAILBOXQUERY, erfolgt die Abfrage der Mailboxen durch den Benutzer."}
    # Alle verfügbaren Postfächer einsammeln
    $MAILBOXES = for ($i=1; $i -le $ns.Folders.Count; $i++) { $ns.Folders.Item($i).Name }

    $quoted = ($MAILBOXES | ForEach-Object { '"{0}"' -f $_ }) -join ',' # Schönere Formatierung
    If ($script:FILELOGGING) {Write-Log "Verfügbare Mailboxen = $quoted"}

    # Alle verfübaren Postfächer auf der Console ausgeben inklusive Zurordnung einer Nummerierung
    If ($script:FILELOGGING) {Write-Host "`nVerfügbare Postfächer:`n"}
    for ($i=0; $i -lt $MAILBOXES.Count; $i++) {
        Write-Host "$($i+1): $($MAILBOXES[$i])"
    }

    # Auf die Auswahl des Benutzers warten
    do {
        $sel = Read-Host "Nummern (max. 3) kommasepariert eingeben"
    } until ($sel -match '^\d+(,\d+){0,2}$')

    # Nimmt die Benutzereingabe und baut daraud ein neues Array mit den gewählten Postfächer
    $MAILBOXES = $sel.Split(',') | ForEach-Object{
        $MAILBOXES[[int]($_)-1]
    }
    # $MAILBOXES enthält jetzt nur die ausgewählten Postfächer

    $quoted = ($MAILBOXES | ForEach-Object { '"{0}"' -f $_ }) -join ',' # Schönere Formatierung
    If ($script:FILELOGGING) {Write-Log "Gewählte Mailbox(en) = $quoted"}
}

# ──────────────────────────Vorhandensein der Excel-Vorlage prüfen─────────────────────────────
if (-not (Test-Path $EXCELTEMPLATE)) {
    Write-Error "Vorlagendatei nicht gefunden: $EXCELTEMPLATE"
    exit
}
else {
    If ($script:FILELOGGING) {Write-Log "Vorlagendatei gefunden: $EXCELTEMPLATE"}
}

# ────────────────────────────────Datumspanne bestimmen────────────────────────────────────────

# Wenn noch StartDate definiert wurde, dann erfolgt Auswertung der Attribute YearsBack und MonthsBack

if (-not $STARTDATE) {
    If ($script:FILELOGGING) {Write-Log "Berechnung Datumsspanne basiert auf aktuelles Datum abzüglich der Parameter YEARSBACK = $YEARSBACK und MONTHBACK = $MONTHBACK."}

    $STARTDATE = if ($YEARSBACK -eq 0 -and $MONTHBACK -eq 0) {
                    # Wenn YEARSBACK und MONTHBACK gleich 0, dann Minimal-Datum setzen
                    [DateTime]::MinValue				 
                } else {
                    # Wenn YEARSBACK und/oder MONTHBACK ungleich 0, dann Startdatum berechnen
                    (Get-Date).AddYears(-$YEARSBACK).AddMonths(-$MONTHBACK)
                }
    If ($script:FILELOGGING) {Write-Log "STARTDATE = 'Aktuelles Datum' - YEARSBACK - MONTHBACK = $STARTDATE"}  
    
    $STARTDATE = $STARTDATE.Date # Setze Uhrzeit auf 00:00:00
    If ($script:FILELOGGING) {Write-Log "Setze Uhrzeit auf 00:00:00: STARTDATE = $STARTDATE"}              
    
    $ENDDATE = Get-Date
    If ($script:FILELOGGING) {Write-Log "ENDDATE = $ENDDATE"}
    $ENDDATE = $ENDDATE.Date # Setze Uhrzeit auf 00:00:00
    $ENDDATE = $ENDDATE.AddDays(1).AddSeconds(-1) # Setze Uhrzeit auf 23:59:59
    If ($script:FILELOGGING) {Write-Log "Setze Uhrzeit aug 23:59:59: ENDDATE = $ENDDATE"}
}
else {
    # Attribut StartDate liegt vor.
    If ($script:FILELOGGING) {Write-Log "Berechnung Datumsspanne basiert auf den Parametern STARTDATE und ENDDATE."}
    try {
        if ($STARTDATE) {
            $STARTDATE = [datetime]::ParseExact($STARTDATE, 'yyyy-MM-dd', $null)
            If ($script:FILELOGGING) {Write-Log "STARTDATE = $STARTDATE"}
        } else {
            $STARTDATE = [DateTime]::MinValue # Wenn kein Parsen möglich, dann wird StartDate auf das frühest mögliche Datum gesetzt
            If ($script:FILELOGGING) {Write-Log "STARTDATE nicht verfügbar, daher auf Minimum gesetzt: $STARTDATE"}
        }
        $STARTDATE = $STARTDATE.Date # Setze Uhrzeit auf 00:00:00
        If ($script:FILELOGGING) {Write-Log "STARTDATE = $STARTDATE"}

        # Dann wird geprüft, ob das Attribut EndDate im Format yyyy-mm-dd übergeben wurde.
        if ($ENDDATE) {
            $ENDDATE = [datetime]::ParseExact($ENDDATE, 'yyyy-MM-dd', $null)
            If ($script:FILELOGGING) {Write-Log "ENDDATE = $ENDDATE"}
        } else {
            $ENDDATE = Get-Date # Wenn kein Parsen möglich, dann wird EndDate auf das aktuelle Datum gesetzt
            If ($script:FILELOGGING) {Write-Log "ENDDATE nicht verfügbar, daher auf Minimum gesetzt: $ENDDATE"}
        }
        $ENDDATE = $ENDDATE.Date # Setze Uhrzeit auf 00:00:00
        $ENDDATE = $ENDDATE.AddDays(1).AddSeconds(-1) # Setze Uhrzeit auf 23:59:59
        If ($script:FILELOGGING) {Write-Log "ENDDATE = $ENDDATE"}

    } catch {
        Write-Host "FEHLER: Ungültiges Datumsformat. Bitte verwende 'yyyy-MM-dd'." -ForegroundColor Red
        Write-Error  "FEHLER: Ungültiges Datumsformat. Bitte verwende 'yyyy-MM-dd'."
        exit
    }
}

Write-Host "Ich exportiere die Emails ab: $STARTDATE bis $ENDDATE" # Ausgabe zur Kontrolle
If ($script:FILELOGGING) {Write-Log "Ich exportiere die Emails ab: $STARTDATE bis $ENDDATE"}

# ────────────────────────────────────Mails einsammeln─────────────────────────────────────────

# Lösche zur Sicherheit die Variablen $script:stats und $seen aus dem aktuellen Scope
Remove-Variable -Name stats,seen -ErrorAction SilentlyContinue

# Erzeugung eine dynamisch wachsende ArrayList zur Speicherung von E-Mail-Statistikdaten (jeweils als Objekt)
#$script:stats = [System.Collections.ArrayList]::new() 

# Erzeugung eines HashSet für Strings zur Duplikatserkennung
$seen  = New-Object 'System.Collections.Generic.HashSet[string]'

# Liste der nicht relevanten Ordner
$script:skipFolders = @(
    'Kontakte','Contacts',
    'Kalender','Calendar',
    'Aufgaben','Tasks',
    'Journal',
    'Notes',
    'Deleted Items','Gelöschte Objekte','Gelöschte Elemente',
    'Yammer',
    'Kalender & Abwesenheiten',
    'Entwürfe','Drafts',
    'Trash',
    'Synchronisierungsprobleme',
    'Metadata'
    ) 

$script:TestCounter = 0
$script:FolderCounter = 0
#$script:ItemCounter = 0

$script:stats = [System.Collections.Generic.List[pscustomobject]]::new()

# Die ausgewählten Postfächer werden nacheinander abgearbeitet
foreach ($Mailbox in $MAILBOXES) {

    If ($script:FILELOGGING) {Write-Log "Versuche auf das Postfach '$Mailbox' zuzugreifen..."}

    # Suche im obersten Ordner-Level nach einem Postfach mit dem Namen $Mailbox.
    # Falls kein solches Postfach gefunden wird ($mbx ist $null), wird mit throw eine Fehlermeldung erzeugt und das Skript bricht kontrolliert ab.
    $mbx = $ns.Folders.Item($Mailbox)  
    if (-not $mbx) { 
        Write-Host "Kein Zugriff auf das Postfach '$Mailbox'."
        Write-Log "Kein Zugriff auf das Postfach '$Mailbox'." -Level WARN
        continue # Weiter zum nächsten Postfach
    }
    else {
        Write-Host "Zugriff auf das Postfach '$Mailbox' ist vorhanden."
        If ($script:FILELOGGING) {Write-Log "Zugriff auf das Postfach '$Mailbox' ist vorhanden."} 
    }

    # Such nach der Inbox des Postfaches
    $inbox = Get-InboxFolder $mbx
    if ($null -eq $inbox) {
        Write-Log "Mailbox '$($mbx.Name)' hat keinen Inbox-Ordner, daher keine Emailsuche möglich." -Level WARN
        continue    # Nächstes Postfach
    }
    If ($script:FILELOGGING) {Write-Log ("Inbox-Ordner '{0}' gefunden." -f $inbox.Name)}

    # Aufruf der Scan-Funktion mit dem Postfach mit dem Namen $Mailbox
    Scan -CurrentMailbox:$Mailbox  -fld $mbx | Out-Null

}

Write-Log ("FolderIgnoreCounter = $script:FolderIgnoreCounter")
Write-Log ("IPMNoteIgnoreCounte = $script:IPMNoteIgnoreCounter")
Write-Log ("NoSentOnIgnoreCounter = $script:NoSentOnIgnoreCounter")
Write-Log ("ItemOlderStartdateCounter = $script:ItemOlderStartdateCounter")

# Abbruch, wenn keine Emails gefunden wurden
if($script:stats.Count -eq 0){Write-Warning 'No mails found.';return}

# ---------- 4  Vorlage kopieren & Daten schreiben ----------------------

# Namen der Excel-Export-Datei erzeugen
$timestamp = Get-Date -Format 'yyyy-MM-dd_HH-mm'
$outFile   = Join-Path $OUTDIR "MailStatistic_$timestamp.xlsm"


try {       

    # Start (bzw. Verbindung mit) Excel über die Windows-COM-Schnittstelle.
    $xl = New-Object -ComObject Excel.Application

    # Excel im Hintergrund laufen lassen (nicht sichtbar für den Benutzer).
    # Das beschleunigt das Skript und vermeidet visuelle Störungen.
    $xl.Visible = $false

    # Öffnet die Excel-Datei, die als Vorlage dient (z. B. mit vordefiniertem Makro und Formatierung).
    # Parameter:
    # 1. $EXCELTEMPLATE – Pfad zur Vorlagendatei
    # 2. $null – keine speziellen Update-Einstellungen
    # 3. $true – Datei wird im **Nur-Lesen-Modus** geöffnet (zum Schutz der Vorlage)
    $wb = $xl.Workbooks.Open($EXCELTEMPLATE,$null,$true)

    # Speichert eine **Kopie** der geöffneten Vorlage unter dem Namen `$outFile`.
    # Das bedeutet: Die Vorlage selbst bleibt unverändert, es wird nur mit einer Kopie gearbeitet.
    $wb.SaveCopyAs($outFile)

    # Schließt die geöffnete Vorlage (die im Nur-Lesen-Modus geöffnet war), ohne sie zu speichern.
    #$wb.Close($false)

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

    foreach ($entry in $script:stats) {
        $ws.Cells($row,1).Value = $entry.StoreID
        $ws.Cells($row,2).Value = $entry.EntryID

        $cell = $ws.Cells($row, 3)
        $cell.Value = "Open"
        $cell.Font.Color = 16711680   # Blau
        $cell.Font.Underline = 2

        $ws.Cells($row,4).Value = Format-DateString $entry.SentOn
        $ws.Cells($row,5).Value = $entry.Sender
        $ws.Cells($row,6).Value = $entry.BehalfOf
        $ws.Cells($row,7).Value = $entry.Subject
        $ws.Cells($row,8).Value = $entry.Folder
        $ws.Cells($row,9).Value = "$($entry.Words)"
        $ws.Cells($row,10).Value = $entry.Recipients

        # Einen zusätzlichen Vergleichsschlüssel erzeugen, um unerkannte Dobletten zu eleminieren
        $comparisonKey = (
            '{0}|{1}|{2}' -f
            (($entry.Subject            -as [string]) ?? '').ToUpper(),
            (($entry.SenderEmailAddress -as [string]) ?? '').ToUpper(),
            (Format-DateString $entry.SentOn)
        )
        $ws.Cells($row, 11).Value2 = $comparisonKey
        $ws.Cells($row,13).Value = $entry.Mailbox

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
            $script:DoublettenCounter++
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
    #$wb2.Close($false) | Out-Null 
    #$xl.Quit() | Out-Null

}
finally {
    # Reihenfolge umgekehrt zur Erzeugung!
    $null = $sortRange, $sort, $ws |
            ForEach-Object {
                if ($_){ [void][Runtime.InteropServices.Marshal]::ReleaseComObject($_) }
            }

    if ($wb2)  { $wb2.Close($false);  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb2) }
    if ($wb)   { $wb.Close($false);   [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)  }

    if ($xl)   { $xl.Quit();          [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) }

    # Alle Variablen kappen, dann Garbage Collector anstoßen
    $sortRange = $sort = $ws = $wb2 = $wb = $xl = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

# COM-Objekt manuell aus dem Speicher entladen, um Ressourcen freizugeben.
[Runtime.InteropServices.Marshal]::ReleaseComObject($ol)|Out-Null

Write-Log ("DoublettenCounter = $script:DoublettenCounter")

Write-Host "Done -> $outFile  ($($script:stats.Count) rows)"
Write-Host "Hinweis: Damit der Link auf die Emails funktioniert muss beim ersten Öffnen von $outName 'Inhalt aktivieren' bestätigt werden."

## ------ Outlook-Cleanup (einmalig aufrufen) ------
try {
    # 1) Outlook sanft schließen – wenn möglich
    try { $script:ol.Quit() } catch { }

    # 2) RCW freigeben – egal, ob noch verbunden oder schon getrennt
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($script:ol)
}
catch {
    Write-Warning "RCW war bereits getrennt: $_"
}
finally {
    # 3) Variable kappen und Speicher aufräumen
    $script:ol = $null
    Remove-Variable -Name ol -Scope Script -ErrorAction SilentlyContinue
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

# Am Ende - Bei PSDebug müssen die beiden folgenden Zeilen aktiviert werden
# Set-PSDebug -Off
# Stop-Transcript

# SIG # Begin signature block
# MIInWQYJKoZIhvcNAQcCoIInSjCCJ0YCAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+DdpXSUIzYpSyIHpbmJcd3y4
# Y3aggiA4MIIFyTCCBLGgAwIBAgIQG7WPJSrfIwBJKMmuPX7tJzANBgkqhkiG9w0B
# AQwFADB+MQswCQYDVQQGEwJQTDEiMCAGA1UEChMZVW5pemV0byBUZWNobm9sb2dp
# ZXMgUy5BLjEnMCUGA1UECxMeQ2VydHVtIENlcnRpZmljYXRpb24gQXV0aG9yaXR5
# MSIwIAYDVQQDExlDZXJ0dW0gVHJ1c3RlZCBOZXR3b3JrIENBMB4XDTIxMDUzMTA2
# NDMwNloXDTI5MDkxNzA2NDMwNlowgYAxCzAJBgNVBAYTAlBMMSIwIAYDVQQKExlV
# bml6ZXRvIFRlY2hub2xvZ2llcyBTLkEuMScwJQYDVQQLEx5DZXJ0dW0gQ2VydGlm
# aWNhdGlvbiBBdXRob3JpdHkxJDAiBgNVBAMTG0NlcnR1bSBUcnVzdGVkIE5ldHdv
# cmsgQ0EgMjCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCCAgoCggIBAL35ePjm1YAM
# ZJ2GG5ZkZz8iOh51AX3v+1xnjMnMXGupkea5QuUgS5vam3u5mV3Zm4BL14RAKyfT
# 6Lowuz4JGqdJle8rQCTCl8en7psl76gKAJeFWqqd3CnJ4jUH63BNStbBs1a4oUE4
# m9H7MX+P4F/hsT8PjhZJYNcGjRj5qiYQqyrT0NFnjRtGvkcw1S5y0cVj2udjeUR+
# S2MkiYYuND8pTFKLKqfA4pEoibnAW/kd2ecnrf+aApfBxlCSmwIsvam5NFkKv4RK
# /9/+s5/r2Z7gmCPspmt3FirbzK07HKSH3EZzXhliaEVX5JCCQrtC1vBh4MGjPWaj
# XfQY7ojJjRdFKZkydQIx7ikmyGsC5rViRX83FVojaInUPt5OJ7DwQAy8TRfLTaKz
# HtAGWt32k89XdZn1+oYaZ3izv5b+NNy951JW5bPldXvXQZEF3F1p45UNQ7n8g5Y5
# lXtsgFpPE3LG130pekS6UqQq1UFGCSD+IqC2WzCNvIkM1ddw+IdS/drvrFEuB7NO
# /tAJ2nDvmPpW5m3btVdL3OUsJRXIni54TvjanJ6GLMpX8xrlyJKLGoKWesO8UBJp
# 2A5aRos66yb6I8m2sIG+QgCk+Nb+MC7H0kb25Y51/fLMudCHW8wGEGC7gzW3Xmfe
# R+yZSPGkoRX+rYxijjlVTzkWubFjnf+3AgMBAAGjggE+MIIBOjAPBgNVHRMBAf8E
# BTADAQH/MB0GA1UdDgQWBBS2oVQ5AsOgP46KvPrU+Bym0ToO/TAfBgNVHSMEGDAW
# gBQIds3LB/8k9sXN7buQvOKEN0Z19zAOBgNVHQ8BAf8EBAMCAQYwLwYDVR0fBCgw
# JjAkoCKgIIYeaHR0cDovL2NybC5jZXJ0dW0ucGwvY3RuY2EuY3JsMGsGCCsGAQUF
# BwEBBF8wXTAoBggrBgEFBQcwAYYcaHR0cDovL3N1YmNhLm9jc3AtY2VydHVtLmNv
# bTAxBggrBgEFBQcwAoYlaHR0cDovL3JlcG9zaXRvcnkuY2VydHVtLnBsL2N0bmNh
# LmNlcjA5BgNVHSAEMjAwMC4GBFUdIAAwJjAkBggrBgEFBQcCARYYaHR0cDovL3d3
# dy5jZXJ0dW0ucGwvQ1BTMA0GCSqGSIb3DQEBDAUAA4IBAQBRwqFYFiIQi/yGMdTC
# MtNc+EuiL2o+TfirCB7t1ej65wgN7LfGHg6ydQV6sQv613RqAAYfpM6q8mt92BHA
# EQjUDk1hxTqo+rHh45jq4mP9QfWTfQ28XZI7kZplutBfTL5MjWgDEBbV8dAEioUz
# +TfnWy4maUI8us281HrpTZ3a50P7Y1KAhQTEJZVV8H6nnwHFWyj44M6GcKYnOzn7
# OC6YU2UidS3X9t0iIpGW691o7T+jGZfTOyWI7DYSPal+zgKNBZqSpyduRbKcYoY3
# DaQzjteoTtBKF0NMxfGnbNIeWGwUUX6KVKH27595el2BmhaQD+G78UoA+fndvu2q
# 7M4KMIIGZjCCBE6gAwIBAgIQexogTp0PEAcqD2zdS7QiEDANBgkqhkiG9w0BAQsF
# ADBWMQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBT
# LkEuMSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBTaWduaW5nIDIwMjEgQ0EwHhcNMjUw
# NTE2MTg1MzE4WhcNMjYwNTE2MTg1MzE3WjCBizELMAkGA1UEBhMCREUxEDAOBgNV
# BAgMB0JhdmFyaWExGTAXBgNVBAcMEE9iZXJzY2hsZWnDn2hlaW0xHjAcBgNVBAoM
# FU9wZW4gU291cmNlIERldmVsb3BlcjEvMC0GA1UEAwwmT3BlbiBTb3VyY2UgRGV2
# ZWxvcGVyLCBSw7xkaWdlciBaw7ZsY2gwggGiMA0GCSqGSIb3DQEBAQUAA4IBjwAw
# ggGKAoIBgQDSeT2q+oiewNUF7DEGoXEnagmuKUzF/Cu0LM/Cmnqbg33GzLb9g35K
# KZh+wmUhmly6gsZfYqBVolC4q1mNS7Em6XEt+KCi6XoeSkWbq4gWsgBhe+m6OMdR
# eG5d7GsT9r+SzQbUX52KH7WN4YnxK0v7r5Uyg0mLoxOzUL1if9sry+bBLHdMvmSU
# oXTM5dR/qdfJyNPvPuam0zfyuetg0ehyGJVu/9E2jXQxKXH1SFn9iazKvzUwIYKV
# O8fP+JSO7kXjXiMVrI7SoBqrSNSLl/0vzNwfe5hXHVEP+UVtoUvDPD2LwKew9+qx
# BZNhdithFTCkPrJ+10b+v7ADQnMHQQ60pbC50Ex9coL0/DsjtJnI7Sq7YJYWiSfA
# xNW1+4DNQIxwPY1L8CYsT6mJJzY3W7+G5/fSwIgYeafJR8ZmMkm+w8UlnSDqNF56
# jkIBctG1MTt0MrRiuBTORls+AXNB/OgFipTGPpeBBBsdHtiap64a1SmqhGbFdGLI
# jieuew8YITcCAwEAAaOCAXgwggF0MAwGA1UdEwEB/wQCMAAwPQYDVR0fBDYwNDAy
# oDCgLoYsaHR0cDovL2Njc2NhMjAyMS5jcmwuY2VydHVtLnBsL2Njc2NhMjAyMS5j
# cmwwcwYIKwYBBQUHAQEEZzBlMCwGCCsGAQUFBzABhiBodHRwOi8vY2NzY2EyMDIx
# Lm9jc3AtY2VydHVtLmNvbTA1BggrBgEFBQcwAoYpaHR0cDovL3JlcG9zaXRvcnku
# Y2VydHVtLnBsL2Njc2NhMjAyMS5jZXIwHwYDVR0jBBgwFoAU3XRdTADbe5+gdMqx
# bvc8wDLAcM0wHQYDVR0OBBYEFBo0i0GU7jSKYVUmZzdqDB4Vll/SMEsGA1UdIARE
# MEIwCAYGZ4EMAQQBMDYGCyqEaAGG9ncCBQEEMCcwJQYIKwYBBQUHAgEWGWh0dHBz
# Oi8vd3d3LmNlcnR1bS5wbC9DUFMwEwYDVR0lBAwwCgYIKwYBBQUHAwMwDgYDVR0P
# AQH/BAQDAgeAMA0GCSqGSIb3DQEBCwUAA4ICAQBauPoakIzbi5QycxReXBK7VHYl
# bq5Vh3M2TeMbXknKi8BhYypzE3z+gsjwdAX+cvhUcg1HVNtwXFzg9ik0IltfL50/
# MFgt67ZHaY0+i0gfJ/6MNwYofVbgQzYqATlhJwGo/ZfptI3yuvldip+9bUsUCPYM
# P5yBLUSXGdMs5hl9lE0js67e3QyuOfilLoogw2WMh9IV+186HwNygbvSs3PehT3E
# Nne0QdssLRytCWoOsIkoaNe98FR0Vk7DS+l3WtLpYB75bOgJmGXaYR0jkB3jKcAB
# DVr92UPPg5aJbf/2VnmH9TGe5AB8wRG7MbWDJNCa9ABVDIOH5PYBcoQfoSRA3h69
# RnxMH5y8a6/HJoSVx8xH6h532QwDQzKlVAn0RY3ypp0zXQx35WJZxxmwXti7Lvya
# S7smMQFSlkkwILsZrRul7FrGxkZalLsFvEo6mH+89BvvlL1WNNaUmqIKWiCQw3TP
# LKw/FKxT6Jj0hXCb6Ch4Mk1RWu9QCaLlPfSbBLwcovcijONMPmwNA3v7FEU5EY8V
# ZXR7vhMwfaIDfHxTPObdwih7sBGd4iEMkq1d9fALGyj7yjQM988ZamF/3yiruQui
# stR+riplOqeOUmFuqKWjIVDa39iu709GDmARdVl4Yqc8FNKwc6g3hpUoS5E66b+V
# MZlz59BcsPfCbEa9GTCCBoMwggRroAMCAQICEQCenAT2Vai0pwJtSYxseI2qMA0G
# CSqGSIb3DQEBDAUAMFYxCzAJBgNVBAYTAlBMMSEwHwYDVQQKExhBc3NlY28gRGF0
# YSBTeXN0ZW1zIFMuQS4xJDAiBgNVBAMTG0NlcnR1bSBUaW1lc3RhbXBpbmcgMjAy
# MSBDQTAeFw0yNTAxMDkwODQwNDNaFw0zNjAxMDcwODQwNDNaMFAxCzAJBgNVBAYT
# AlBMMSEwHwYDVQQKDBhBc3NlY28gRGF0YSBTeXN0ZW1zIFMuQS4xHjAcBgNVBAMM
# FUNlcnR1bSBUaW1lc3RhbXAgMjAyNTCCAiIwDQYJKoZIhvcNAQEBBQADggIPADCC
# AgoCggIBAMcpX2f4rCvdkEXlQItY5D811tuUC9B4YyN8KL9uJtPvWE+TApXhS9C3
# Uy47ChpQJi2wuHDyOuY6ajPYegcorMSmJLcDWsDeyxdRkydKtuKJWkOP7ky2dtdJ
# UQkptxy4dubByOXf03hbvbNxygL+d4oC7A7UMn71I75iQedxGJX3mJ1aHFEBwNi/
# juFz0YQVm1MXFBctusvg3s33QGouV7drLOfNxwQ9D4TofqnTTMT0dSn77dtlMXg9
# /I9GcoItzl7MDMSqpttTkX+e6PHr1PWabTaOW2UVedMwXW+VqR5BdZikYiO2tHtY
# /u0gxmeTvyethJ94CnyV779b/n6qvu2pBm59b2Ke+xFNx+Ts/6acqMku5JvCqDF6
# W8IGvLP2vnVtPTnrVpHZz6OgXKCpbufgU8Kcws6PjRfii3vA/a7yQtENUduGwerl
# IYuqvwNsWrXjCzYAtLe1hrYYaIDvBfv2tLmgMydJj/GYk7ly6m5K+qnKXM7uPNws
# 7BaKoKurTkBennVJGuP3FG3wp1lakN1aFaYTwwWIW2ZDQCo4MskGfWFADHcW8P+m
# 17qkTUGKXQE5ULZKdt6r/c3KUPC8N0CD/QJlEmxAhrJ0ezAnDJgGVhQPJpArbJRL
# nuGhudjH+9K0y9MH3qKLiMZXdSRmHca/B0MKxmRsldpTwBMX/PoPAgMBAAGjggFQ
# MIIBTDB1BggrBgEFBQcBAQRpMGcwOwYIKwYBBQUHMAKGL2h0dHA6Ly9zdWJjYS5y
# ZXBvc2l0b3J5LmNlcnR1bS5wbC9jdHNjYTIwMjEuY2VyMCgGCCsGAQUFBzABhhxo
# dHRwOi8vc3ViY2Eub2NzcC1jZXJ0dW0uY29tMB8GA1UdIwQYMBaAFL5UAi+/QGxz
# Q86sCSVOnkNEGu7gMAwGA1UdEwEB/wQCMAAwOQYDVR0fBDIwMDAuoCygKoYoaHR0
# cDovL3N1YmNhLmNybC5jZXJ0dW0ucGwvY3RzY2EyMDIxLmNybDAWBgNVHSUBAf8E
# DDAKBggrBgEFBQcDCDAOBgNVHQ8BAf8EBAMCB4AwIgYDVR0gBBswGTAIBgZngQwB
# BAIwDQYLKoRoAYb2dwIFAQswHQYDVR0OBBYEFIGMBqAoU/wAP9t+nkrBHyLsshKJ
# MA0GCSqGSIb3DQEBDAUAA4ICAQCZDxkMG+sFQ+dovzsBhzgmWZ8fVk/nK3rky3Ni
# 7x9uawvd3dy9iz4Sdu1+71/bAx6HNLK953dFzn1rTg7w03umXDg9eDvXB9ITgxEU
# zUS6ovrfOr25OOb/20DvevVoJ1aaSmsqnRouhmiVQ3SoZy+v35AbxUglUxy5KoV4
# S9GssQorbFWQxJ1NNsOGx5SMji834GtebnPjkQDdoOlKJlC6g13hEWcPN11uB9wJ
# A9pjZTJerM0GOe7PoDIecXMq02UJ6+QwGCHh0gO3/QMYYM5pMQBZm1QSorkCUd3Q
# 8Gd3jQueiDhQBNTcG3oYd94OZYQcVOMcqyaf1DzCIaP3TptWAvzfm18Qf0SBgOSW
# qs5TFbEN/Vw4Dt2z+vykhza+MtD205KSb1ZdVudN4sbdDnt+tOAK9M5t3+p/dTMT
# 3udM05Tu84xkSjjUCCaG4RsazJAgMz9Xp8lBfVGfPMk2ip9NORTNSg/a2U0ec2yM
# WjZhX7nJlyhprCY1aZHtPLBRcbb+8WAlobRt8Sih155ensDgjgdoMOl9FrvKmAkQ
# 95PAe4FEmmP037XG2uL7oHMc/CwAC9Qmbnw8ahWy14cBfC+mDg1WC9STcEpuXvEW
# 6VZTUoeof+yTtPQIUFDmiG4o+YV/WioA/gA71rdiDbAEzASDXM86HrFtReyTOz2J
# Jr3IbDCCBrkwggShoAMCAQICEQCZo4AKJlU7ZavcboSms+o5MA0GCSqGSIb3DQEB
# DAUAMIGAMQswCQYDVQQGEwJQTDEiMCAGA1UEChMZVW5pemV0byBUZWNobm9sb2dp
# ZXMgUy5BLjEnMCUGA1UECxMeQ2VydHVtIENlcnRpZmljYXRpb24gQXV0aG9yaXR5
# MSQwIgYDVQQDExtDZXJ0dW0gVHJ1c3RlZCBOZXR3b3JrIENBIDIwHhcNMjEwNTE5
# MDUzMjE4WhcNMzYwNTE4MDUzMjE4WjBWMQswCQYDVQQGEwJQTDEhMB8GA1UEChMY
# QXNzZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQwIgYDVQQDExtDZXJ0dW0gQ29kZSBT
# aWduaW5nIDIwMjEgQ0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQCd
# I88EMCM7wUYs5zNzPmNdenW6vlxNur3rLfi+5OZ+U3iZIB+AspO+CC/bj+taJUbM
# bFP1gQBJUzDUCPx7BNLgid1TyztVLn52NKgxxu8gpyTr6EjWyGzKU/gnIu+bHAse
# 1LCitX3CaOE13rbuHbtrxF2tPU8f253QgX6eO8yTbGps1Mg+yda3DcTsOYOhSYNC
# JiL+5wnjZ9weoGRtvFgMHtJg6i671OPXIciiHO4Lwo2p9xh/tnj+JmCQEn5QU0Nx
# zrOiRna4kjFaA9ZcwSaG7WAxeC/xoZSxF1oK1UPZtKVt+yrsGKqWONoK6f5EmBOA
# VEK2y4ATDSkb34UD7JA32f+Rm0wsr5ajzftDhA5mBipVZDjHpwzv8bTKzCDUSUuU
# mPo1govD0RwFcTtMXcfJtm1i+P2UNXadPyYVKRxKQATHN3imsfBiNRdN5kiVVeqP
# 55piqgxOkyt+HkwIA4gbmSc3hD8ke66t9MjlcNg73rZZlrLHsAIV/nJ0mmgSjBI/
# TthoGJDydekOQ2tQD2Dup/+sKQptalDlui59SerVSJg8gAeV7N/ia4mrGoiez+Sq
# V3olVfxyLFt3o/OQOnBmjhKUANoKLYlKmUpKEFI0PfoT8Q1W/y6s9LTI6ekbi0ig
# EbFUIBE8KDUGfIwnisEkBw5KcBZ3XwnHmfznwlKo8QIDAQABo4IBVTCCAVEwDwYD
# VR0TAQH/BAUwAwEB/zAdBgNVHQ4EFgQU3XRdTADbe5+gdMqxbvc8wDLAcM0wHwYD
# VR0jBBgwFoAUtqFUOQLDoD+Oirz61PgcptE6Dv0wDgYDVR0PAQH/BAQDAgEGMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMDAGA1UdHwQpMCcwJaAjoCGGH2h0dHA6Ly9jcmwu
# Y2VydHVtLnBsL2N0bmNhMi5jcmwwbAYIKwYBBQUHAQEEYDBeMCgGCCsGAQUFBzAB
# hhxodHRwOi8vc3ViY2Eub2NzcC1jZXJ0dW0uY29tMDIGCCsGAQUFBzAChiZodHRw
# Oi8vcmVwb3NpdG9yeS5jZXJ0dW0ucGwvY3RuY2EyLmNlcjA5BgNVHSAEMjAwMC4G
# BFUdIAAwJjAkBggrBgEFBQcCARYYaHR0cDovL3d3dy5jZXJ0dW0ucGwvQ1BTMA0G
# CSqGSIb3DQEBDAUAA4ICAQB1iFgP5Y9QKJpTnxDsQ/z0O23JmoZifZdEOEmQvo/7
# 9PQg9nLF/GJe6ZiUBEyDBHMtFRK0mXj3Qv3gL0sYXe+PPMfwmreJHvgFGWQ7Xwnf
# Mh2YIpBrkvJnjwh8gIlNlUl4KENTK5DLqsYPEtRQCw7R6p4s2EtWyDDr/M58iY2U
# BEqfUU/ujR9NuPyKk0bEcEi62JGxauFYzZ/yld13fHaZskIoq2XazjaD0pQkcQiI
# ueL0HKiohS6XgZuUtCKA7S6CHttZEsObQJ1j2s0urIDdqF7xaXFVaTHKtAuMfwi0
# jXtF3JJphrJfc+FFILgCbX/uYBPBlbBIP4Ht4xxk2GmfzMn7oxPITpigQFJFWuzT
# MUUgdRHTxaTSKRJ/6Uh7ki/pFjf9sUASWgxT69QF9Ki4JF5nBIujxZ2sOU9e1HSC
# JwOfK07t5nnzbs1LbHuAIGJsRJiQ6HX/DW1XFOlXY1rc9HufFhWU+7Uk+hFkJsfz
# qBz3pRO+5aI6u5abI4Qws4YaeJH7H7M8X/YNoaArZbV4Ql+jarKsE0+8XvC4DJB+
# IVcvC9Ydqahi09mjQse4fxfef0L7E3hho2O3bLDM6v60rIRUCi2fJT2/IRU5ohgy
# Tch4GuYWefSBsp5NPJh4QRTP9DC3gc5QEKtbrTY0Ka87Web7/zScvLmvQBm8JDFp
# DjCCBrkwggShoAMCAQICEQDn/2nHOzXOS5Em2HR8aKWHMA0GCSqGSIb3DQEBDAUA
# MIGAMQswCQYDVQQGEwJQTDEiMCAGA1UEChMZVW5pemV0byBUZWNobm9sb2dpZXMg
# Uy5BLjEnMCUGA1UECxMeQ2VydHVtIENlcnRpZmljYXRpb24gQXV0aG9yaXR5MSQw
# IgYDVQQDExtDZXJ0dW0gVHJ1c3RlZCBOZXR3b3JrIENBIDIwHhcNMjEwNTE5MDUz
# MjA3WhcNMzYwNTE4MDUzMjA3WjBWMQswCQYDVQQGEwJQTDEhMB8GA1UEChMYQXNz
# ZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQwIgYDVQQDExtDZXJ0dW0gVGltZXN0YW1w
# aW5nIDIwMjEgQ0EwggIiMA0GCSqGSIb3DQEBAQUAA4ICDwAwggIKAoICAQDpEh8E
# Ne25XXrFppVBvoplf0530W0lddNmjtv4YSh/f7eDQKFaIqc7tHj7ox+u8vIsJZlr
# oakUeMS3i3T8aJRC+eQs4FF0GqvkM6+WZO8kmzZfxmZaBYmMLs8FktgFYCzywmXe
# Q1fEExflee2OpbHVk665eXRHjH7MYZIzNnjl2m8Hy8ulB9mR8wL/W0v0pjKNT6G0
# sfrx1kk+3OGosFUb7yWNnVkWKU4qSxLv16kJ6oVJ4BSbZ4xMak6JLeB8szrK9vwG
# DpvGDnKCUMYL3NuviwH1x4gZG0JAXU3x2pOAz91JWKJSAmRy/l0s0l5bEYKolg+D
# MqVhlOANd8Yh5mkQWaMEvBRE/kAGzIqgWhwzN2OsKIVtO8mf5sPWSrvyplSABAYa
# 13rMYnzwfg08nljZHghquCJYCa/xHK9acev9UD7Y+usr15d7mrszzxhF1JOr1Mpu
# p2chNSBlyOObhlSO16rwrffVrg/SzaKfSndS5swRhr8bnDqNJY9TNyEYvBYpgF95
# K7p0g4LguR4A++Z1nFIHWVY5v0fNVZmgzxD9uVo/gta3onGOQj3JCxgYx0KrCXu4
# yc9QiVwTFLWbNdHFSjBCt5/8Q9pLuRhVocdCunhcHudMS1CGQ/Rn0+7P+fzMgWdR
# KfEOh/hjLrnQ8BdJiYrZNxvIOhM2aa3zEDHNwwIDAQABo4IBVTCCAVEwDwYDVR0T
# AQH/BAUwAwEB/zAdBgNVHQ4EFgQUvlQCL79AbHNDzqwJJU6eQ0Qa7uAwHwYDVR0j
# BBgwFoAUtqFUOQLDoD+Oirz61PgcptE6Dv0wDgYDVR0PAQH/BAQDAgEGMBMGA1Ud
# JQQMMAoGCCsGAQUFBwMIMDAGA1UdHwQpMCcwJaAjoCGGH2h0dHA6Ly9jcmwuY2Vy
# dHVtLnBsL2N0bmNhMi5jcmwwbAYIKwYBBQUHAQEEYDBeMCgGCCsGAQUFBzABhhxo
# dHRwOi8vc3ViY2Eub2NzcC1jZXJ0dW0uY29tMDIGCCsGAQUFBzAChiZodHRwOi8v
# cmVwb3NpdG9yeS5jZXJ0dW0ucGwvY3RuY2EyLmNlcjA5BgNVHSAEMjAwMC4GBFUd
# IAAwJjAkBggrBgEFBQcCARYYaHR0cDovL3d3dy5jZXJ0dW0ucGwvQ1BTMA0GCSqG
# SIb3DQEBDAUAA4ICAQC4k1l3yUwV/ZQHCKCneqAs8EGTnwEUJLdDpokN/dMhKjK0
# rR5qX8nIIHzxpQR3TAw2IRw1Uxsr2PliG3bCFqSdQTUbfaTq6V3vBzEebDru9QFj
# qlKnxCF2h1jhLNFFplbPJiW+JSnJTh1fKEqEdKdxgl9rVTvlxfEJ7exOn25MGbd/
# wGPwuSmMxRJVO0wnqgS7kmoJjNF9zqeehFSDDP8ZVkWg4EZ2tIS0M3uZmByRr+1L
# kwjjt8AtW83mVnZTyTsOb+FNfwJY7DS4FmWhkRbgcHRetreoTirPOr/ozyDKhT8M
# TSTf6Lttg6s6T/u08mDWw6HK04ZRDfQ9sb77QV8mKgO44WGP31vXnVKoWVJpFBjP
# vjL8/Zck/5wXX2iqjOaLStFOR/IQki+Ehn4zlcgVm22ZVCBPF+l8nAwUUShCtKuS
# U7GmZLKCmmxQMkSiWILTm8EtVD6AxnJhoq8EnhjEEyUoflkeRF2WhFiVQOmWTwZR
# r44IxWGkNJC6tTorW5rl2Zl+2e9JLPYf3pStAPMDoPKIjVXd6NW2+fZrNUBeDo2e
# Oa5Fn7Brs/HLQff5Xgris5MeUbdVgDrF8uxO6cLPvZPo63j62SsNg55pTWk9fUIF
# 9iPoRbb4QurjoY/woI1RAOKtYtTic6aAJq3u83RIPpGXBSJKwx4KJAOZnCDCtTGC
# BoswggaHAgEBMGowVjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRh
# IFN5c3RlbXMgUy5BLjEkMCIGA1UEAxMbQ2VydHVtIENvZGUgU2lnbmluZyAyMDIx
# IENBAhB7GiBOnQ8QByoPbN1LtCIQMAkGBSsOAwIaBQCgcDAQBgorBgEEAYI3AgEM
# MQIwADAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgorBgEEAYI3AgELMQ4w
# DAYKKwYBBAGCNwIBFTAjBgkqhkiG9w0BCQQxFgQU/OapLV5/5rdBReBLcaHTWqBG
# zUowDQYJKoZIhvcNAQEBBQAEggGAwhF8zeuiFnDiafPH521YuPkR0qf6Ggk6T/cZ
# GeOBtZf7nLqzeJQ55C7mgMNQOwVxKJ/uWLiAjw/76/4rEAZ2RxNRb2kD2ij4n/eF
# VkT4Kb6gixfiSUrfxSyt+skW+vDWV95g0gYxvWWeZfBWT4rJ6AKz/lehvxkPK2el
# 1Ne0oL2cNtRGJFkQJmGMtyxNWOI9VyYIwbYZ+mLX3kn4dpU+lqWnQLo/hna6mdLM
# h8QT/a613W2VtTUmDYk3poUW1PXV4lNaeKQm8uszc4rxoq0plEno/6xFHyMahRQ0
# zsnoPy/REA7VFQhWtZHjQSGeM74CRkvUKvGD0q23mvxvlj0xjFElHVFlcvVtvCCD
# nAg+luW0eYg/e2PttS4Wp8GFBdeRMKNam7baPJXUUSucrOfE1facZcq664V5O4hc
# xlZ8p6yMTTC4Ll7COTXGAAtyndUcZdjINuXbZUhzHc6RMwSSl5SGi1yztmjNRpHi
# 6R4HDvCZ94fkx4vzkbMlEd4ihcdLoYIEBDCCBAAGCSqGSIb3DQEJBjGCA/EwggPt
# AgEBMGswVjELMAkGA1UEBhMCUEwxITAfBgNVBAoTGEFzc2VjbyBEYXRhIFN5c3Rl
# bXMgUy5BLjEkMCIGA1UEAxMbQ2VydHVtIFRpbWVzdGFtcGluZyAyMDIxIENBAhEA
# npwE9lWotKcCbUmMbHiNqjANBglghkgBZQMEAgIFAKCCAVcwGgYJKoZIhvcNAQkD
# MQ0GCyqGSIb3DQEJEAEEMBwGCSqGSIb3DQEJBTEPFw0yNTA2MjgxMDE2MDZaMDcG
# CyqGSIb3DQEJEAIvMSgwJjAkMCIEIM+h3DWd7SvDy4kPojDl2vd7VA8abisj3c8X
# VOGM+qDVMD8GCSqGSIb3DQEJBDEyBDDKI104EyYa7zaSk8dJc6y+J+ixJwhZ82sW
# 2Q1TKIYfBjH9xmESAkgzvEau4gYCkwAwgaAGCyqGSIb3DQEJEAIMMYGQMIGNMIGK
# MIGHBBTDJbibF/zFAmBhzitxe0UH3ZxqajBvMFqkWDBWMQswCQYDVQQGEwJQTDEh
# MB8GA1UEChMYQXNzZWNvIERhdGEgU3lzdGVtcyBTLkEuMSQwIgYDVQQDExtDZXJ0
# dW0gVGltZXN0YW1waW5nIDIwMjEgQ0ECEQCenAT2Vai0pwJtSYxseI2qMA0GCSqG
# SIb3DQEBAQUABIICAHfQtf38lSO4RB9FSPL8M+2WxUrwsdRmdi611neBggWOz4bO
# 9pJfv6G/SNOqEWSgYfYoHqFD8MGRNJ8vhZ9HwumFiIEbDKx6AlnlMnX1IZo6H2KW
# oda0oJURWqogzai8ePTxClSJWhpGjpb7IXbzaXgN4Wx67jjrZsWtI/AgP0VOkea4
# ki+QlMz9MMPjuGnM7oU7KPE5wLuPMUJP6UKRaycKYzo8h+gQ197/9W5gfKtnPw8D
# rrohIaJJBHtK2BCh/h4Z/f3oM3xqDgr4x0jR+f1D4R/3Od2+R9Pxmfe+LJiFmOaU
# jBL00zeIIDEIKpqWUzzxEgwrvLssUYcWaZH+TVH2I3LGJR8QXIcos76AsRE1FjZp
# k/7rjos0toUU5uFpgwCqY+G75qCrJc0nAyvIY8+Ka2loCWUpUt5kBVyDqP2M8sAe
# p7lHPOvMtjgoI4kSyzJDNb/ZQRq1bWIJjn2mJO+zfQIf/7yHluXu7vLUo5fZIvLr
# tMcB1VrwuiBdIN2+soIqDBKbC/1Oh9S5xF04xql4F+A8JpBBLCvTryR/hB8cpqZw
# YiFqM/vSFH8jFPtnkQoRifAozttltcPcX181HFf49k+9o6KC3OXPf5+6DmYFd2Go
# hQDTG9QkRVEvYv38FuJYAY6aPbR1wmrfVpUUBPk8ceTUGSlgshbBY0Y885Nu
# SIG # End signature block
