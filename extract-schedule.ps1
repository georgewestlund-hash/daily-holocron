<#
  extract-schedule.ps1

  Downloads the "8 ECP Literature - Classwork/Homework" Google Slides deck as .pptx,
  parses every cycle slide, joins the rotation grid to the CW/HW grid, and emits a
  normalized schedule as JSON.

  Join model
  ----------
  Each content slide holds two tables:
    * rotation grid : row1 = dates, row2 = rotation label, rows 3+ = one row per period,
                      cells hold "Lesson N" where that period meets.
    * CW/HW grid    : columns = Lesson 1..N, rows = CW/Classwork/Class and HW/Homework.
  So (period, date) -> rotation day -> lesson number -> CW/HW content.

  CW cell decomposition (per teacher's spec)
  ------------------------------------------
    objective : paragraphs before the "Lesson:" marker
    lesson    : paragraphs between "Lesson:" and "Study:"
    study     : paragraphs after "Study:"
    homework  : all paragraphs of the HW row cell
  Cycle 1 uses an unmarked format (no "Lesson:"/"Study:" labels). Fallback there:
  first paragraph is the objective, the rest is the lesson, study is empty.

  Usage:  powershell -ExecutionPolicy Bypass -File extract-schedule.ps1
#>

param(
    # Site root. schedule.json is written to <SiteDir>/data/schedule.json.
    [string]$SiteDir = (Join-Path $PSScriptRoot 'docs'),
    [int]$FallSchoolYear = 2026,
    # Override to parse a single deck in isolation (testing). Normally leave unset
    # so all courses in $Courses below are parsed and merged.
    [string]$DeckId,
    [string[]]$DeckPeriods
)

# One entry per course. `periods` selects which rows of that deck's rotation grid
# belong to the course, so the merged output has exactly one record per
# (period, date) across all decks.
$Courses = @(
    @{ id = '1qk_g2tVKqYpposzORBzEtqEe3fsdVhcUfih14d4z46A'; label = '8 Honors Lit'; periods = @('A', 'D', 'G') },
    @{ id = '1VDpCkYyOhH1SllDQPwJxX81AXgIs8XzHlapY4RAvHaU'; label = '8 ECP Lit'; periods = @('E') }
)
if ($DeckId) { $Courses = @(@{ id = $DeckId; label = 'deck'; periods = $DeckPeriods }) }

$ErrorActionPreference = 'Stop'
# Windows PowerShell 5.1 defaults to TLS 1.0; pwsh on Linux (GitHub Actions) does not
# expose ServicePointManager the same way, so this is best-effort.
try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 } catch { }

if (-not $SiteDir) { $SiteDir = Join-Path (Get-Location).Path 'docs' }
$dataDir = Join-Path $SiteDir 'data'
if (-not (Test-Path $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }

$jsonPath = Join-Path $dataDir 'schedule.json'

# Parses one deck end to end and returns its records, cycles, calendar and
# warnings. Defined before the helper functions it calls, which is fine —
# PowerShell resolves function names at call time, not definition time.
function Get-DeckSchedule {
param([string]$DeckId, [string]$DataDir, [int]$FallSchoolYear)

$deckPath = Join-Path $DataDir "deck-$DeckId.pptx"

# ---------------------------------------------------------------- download ---
$exportUrl = "https://docs.google.com/presentation/d/$DeckId/export/pptx"
Write-Host "Fetching $exportUrl"
$tmp = "$deckPath.tmp"
Invoke-WebRequest -Uri $exportUrl -OutFile $tmp -UseBasicParsing -TimeoutSec 120

# Sanity-check: a permission failure returns an HTML error page, not a zip.
$sig = [System.IO.File]::ReadAllBytes($tmp)[0..1]
if ($sig[0] -ne 0x50 -or $sig[1] -ne 0x4B) {
    Remove-Item $tmp -Force
    throw "Export did not return a .pptx (deck may no longer be link-shared). Aborting; previous schedule.json left intact."
}
Move-Item -Path $tmp -Destination $deckPath -Force

# ----------------------------------------------------------------- helpers ---
$BulletChars = @([char]0x2013, [char]0x2014, '-', '#', [char]0x2022, '*')

function Clean-Line {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = $s -replace [char]0x00A0, ' '
    $s = $s.Trim()
    while ($s.Length -gt 0 -and ($BulletChars -contains $s[0])) { $s = $s.Substring(1).Trim() }
    $s = $s -replace '\s{2,}', ' '
    return $s.Trim()
}

function Get-CellParagraphs {
    param($Cell, $Ns)
    $out = @()
    foreach ($p in $Cell.SelectNodes('.//a:p', $Ns)) {
        $raw = (($p.SelectNodes('.//a:t', $Ns) | ForEach-Object { $_.InnerText }) -join '')
        $clean = Clean-Line $raw
        if ($clean -ne '') { $out += $clean }
    }
    return , $out
}

function Get-TableRows {
    param($Tbl, $Ns)
    $rows = @()
    foreach ($tr in $Tbl.SelectNodes('./a:tr', $Ns)) {
        $cells = @()
        foreach ($tc in $tr.SelectNodes('./a:tc', $Ns)) {
            $cells += , (Get-CellParagraphs -Cell $tc -Ns $Ns)
        }
        $rows += , $cells
    }
    return , $rows
}

function Join-Cell { param($Paras) if ($Paras.Count -eq 0) { return '' } return ($Paras -join ' ') }

# Splits a CW cell into objective / lesson / study.
function Split-Classwork {
    param($Paras)
    $obj = @(); $les = @(); $std = @()
    $mode = 'objective'
    $sawLessonMarker = $false
    foreach ($p in $Paras) {
        $probe = ($p -replace '[:\s]+$', '')
        if ($probe -match '^(?i)lesson$') { $mode = 'lesson'; $sawLessonMarker = $true; continue }
        if ($probe -match '^(?i)(study|homework|hw)$') { $mode = 'study'; continue }
        switch ($mode) {
            'objective' { $obj += $p }
            'lesson' { $les += $p }
            'study' { $std += $p }
        }
    }
    if (-not $sawLessonMarker) {
        # Unmarked (Cycle 1) format: first line is the header/objective, rest is the lesson.
        if ($obj.Count -gt 1) {
            $les = $obj[1..($obj.Count - 1)]
            $obj = @($obj[0])
        }
    }
    return [pscustomobject]@{
        objective = (Join-Cell $obj)
        lesson    = [string[]]$les
        study     = [string[]]$std
    }
}

$WeekdayMap = @{
    'mon' = 'Monday'; 'monday' = 'Monday'
    'tue' = 'Tuesday'; 'tues' = 'Tuesday'; 'tuesday' = 'Tuesday'
    'wed' = 'Wednesday'; 'weds' = 'Wednesday'; 'wednesday' = 'Wednesday'
    'thu' = 'Thursday'; 'thur' = 'Thursday'; 'thurs' = 'Thursday'; 'thursday' = 'Thursday'
    'fri' = 'Friday'; 'friday' = 'Friday'
}

function Parse-DateHeader {
    param([string]$Text, [int]$FallYear)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    $m = [regex]::Match($Text, '(\d{1,2})\s*/\s*(\d{1,2})')
    if (-not $m.Success) { return $null }
    $mon = [int]$m.Groups[1].Value
    $day = [int]$m.Groups[2].Value
    $year = $FallYear
    if ($mon -le 7) { $year = $FallYear + 1 }
    $dt = $null
    try { $dt = New-Object DateTime($year, $mon, $day) } catch { return $null }

    $stated = $null
    $wm = [regex]::Match($Text, '^\s*([A-Za-z]+)')
    if ($wm.Success) {
        $key = $wm.Groups[1].Value.ToLower()
        if ($WeekdayMap.ContainsKey($key)) { $stated = $WeekdayMap[$key] }
    }
    return [pscustomobject]@{
        date           = $dt.ToString('yyyy-MM-dd')
        weekday        = $dt.DayOfWeek.ToString()
        statedWeekday  = $stated
        weekdayMismatch = ($null -ne $stated -and $stated -ne $dt.DayOfWeek.ToString())
    }
}

$MonthNames = @('January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December')

# True if the slide's notes line explicitly names this date ("December 10" or "12/10").
function Note-NamesDate {
    param([string]$Note, [string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($Note)) { return $false }
    $d = [datetime]::ParseExact($IsoDate, 'yyyy-MM-dd', $null)
    $long = $MonthNames[$d.Month - 1] + '\s+' + $d.Day + '\b'
    $numeric = '\b' + $d.Month + '\s*/\s*' + $d.Day + '\b'
    return ($Note -match "(?i)$long") -or ($Note -match $numeric)
}

# Maps a deck rotation label to the board's day-type code: 1-7, X, A, B, Bn.
function Get-DayType {
    param([string]$Rotation, [string]$Note, [string]$IsoDate)
    if ([string]::IsNullOrWhiteSpace($Rotation)) { return $null }
    $m = [regex]::Match($Rotation, '(?i)day\s*([1-7])\b')
    if ($m.Success) { return $m.Groups[1].Value }
    if ($Rotation -match '(?i)x\s*day') { return 'X' }
    # The deck sometimes labels an A Block day "B BLOCK (A-D)". The period range
    # is authoritative: A-D means A Block, E-G means B Block.
    if ($Rotation -match '(?i)block') {
        if ($Rotation -match '(?i)\(\s*A\s*-\s*D\s*\)') { return 'A' }
        if ($Rotation -match '(?i)\(\s*E\s*-\s*G\s*\)') {
            if ($Note -match '(?i)no\s+late\s+start' -and (Note-NamesDate -Note $Note -IsoDate $IsoDate)) { return 'Bn' }
            return 'B'
        }
        if ($Rotation -match '(?i)^\s*A\s*BLOCK') { return 'A' }
        if ($Rotation -match '(?i)^\s*B\s*BLOCK') { return 'B' }
    }
    return $null
}

function Normalize-Period {
    param([string]$Text)
    $t = (Clean-Line $Text)
    $t = $t -replace '(?i)\s*period\s*$', ''
    $t = $t.Trim()
    if ($t -match '^(?i)([A-G])$') { return $Matches[1].ToUpper() }
    return $null
}

function Parse-LessonNumber {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    if ($Text -notmatch '(?i)lesson') { return $null }
    $m = [regex]::Match($Text, '(?i)lesson\s*(\d+)')
    if ($m.Success) { return [int]$m.Groups[1].Value }
    return 'UNNUMBERED'
}

# -------------------------------------------------------------------- parse ---
# Needed on Windows PowerShell 5.1; already loaded in pwsh (.NET Core).
try { Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction Stop } catch { }
$zip = [System.IO.Compression.ZipFile]::OpenRead($deckPath)

$slideEntries = $zip.Entries |
    Where-Object { $_.FullName -like 'ppt/slides/slide*.xml' } |
    Sort-Object { [int](($_.FullName -replace '[^0-9]', '')) }

$records = @()
$cycles = @()
$warnings = @()
$calendar = @{}
$deckTitle = ''

foreach ($entry in $slideEntries) {
    $slideNo = [int](($entry.FullName -replace '[^0-9]', ''))
    $reader = New-Object System.IO.StreamReader($entry.Open())
    $xml = [xml]$reader.ReadToEnd()
    $reader.Close()

    $ns = New-Object System.Xml.XmlNamespaceManager($xml.NameTable)
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    $ns.AddNamespace('p', 'http://schemas.openxmlformats.org/presentationml/2006/main')

    # --- slide-level text: title + notes ---
    $shapeTexts = @()
    foreach ($sp in $xml.SelectNodes('//p:sp', $ns)) {
        $t = (($sp.SelectNodes('.//a:t', $ns) | ForEach-Object { $_.InnerText }) -join '')
        $t = Clean-Line $t
        if ($t -ne '') { $shapeTexts += $t }
    }
    $title = ''
    $notes = ''
    foreach ($t in $shapeTexts) {
        if ($t -match '(?i)^comments?/?\s*important\s*notes\s*:?') {
            $notes = Clean-Line ($t -replace '(?i)^comments?/?\s*important\s*notes\s*:?', '')
        }
        elseif ($title -eq '') { $title = $t }
    }
    if ($title -eq '') { continue }   # blank / divider slide
    if ($deckTitle -eq '') { $deckTitle = $title }   # slide 1 is the deck's cover

    # --- classify the two tables ---
    $rotationRows = $null
    $cwRows = $null
    foreach ($tbl in $xml.SelectNodes('//a:tbl', $ns)) {
        $rows = Get-TableRows -Tbl $tbl -Ns $ns
        if ($rows.Count -lt 2) { continue }

        $firstColLabels = @()
        foreach ($r in $rows) { if ($r.Count -gt 0) { $firstColLabels += (Join-Cell $r[0]) } }
        $isCw = ($firstColLabels | Where-Object { $_ -match '^(?i)(cw|classwork|class)$' }).Count -gt 0

        $headerCells = @()
        if ($rows.Count -gt 0) { foreach ($c in $rows[0]) { $headerCells += (Join-Cell $c) } }
        $dateish = ($headerCells | Where-Object { $_ -match '\d{1,2}\s*/\s*\d{1,2}' }).Count

        if ($isCw -and $null -eq $cwRows) { $cwRows = $rows }
        elseif ($dateish -ge 2 -and $null -eq $rotationRows) { $rotationRows = $rows }
    }

    if ($null -eq $rotationRows) {
        # Slide 1 is the cover and legitimately has no grid.
        if ($slideNo -ne 1) { $warnings += "slide${slideNo}: no rotation grid found ('$title')" }
        continue
    }

    # --- cycle metadata ---
    $cycleNo = $null
    $cm = [regex]::Match($title, '(?i)cycle\s*(\d+)')
    if ($cm.Success) { $cycleNo = [int]$cm.Groups[1].Value }
    $isTemplate = ($title -match '(?i)template')
    if ($title -match '(?i)^1st\s+semester' -and $title -notmatch '(?i)template') { }

    # --- lesson content, keyed by lesson number ---
    $lessonContent = @{}
    if ($null -ne $cwRows -and $cwRows.Count -ge 2) {
        $hdr = $cwRows[0]
        $lessonCols = @{}
        for ($c = 1; $c -lt $hdr.Count; $c++) {
            $n = Parse-LessonNumber (Join-Cell $hdr[$c])
            if ($n -is [int]) { $lessonCols[$c] = $n }
        }
        $cwRow = $null; $hwRow = $null
        foreach ($r in $cwRows) {
            if ($r.Count -eq 0) { continue }
            $label = Join-Cell $r[0]
            if ($label -match '^(?i)(cw|classwork|class)$' -and $null -eq $cwRow) { $cwRow = $r }
            elseif ($label -match '^(?i)(hw|homework)$' -and $null -eq $hwRow) { $hwRow = $r }
        }
        foreach ($c in $lessonCols.Keys) {
            $n = $lessonCols[$c]
            $cwParas = @(); $hwParas = @()
            if ($null -ne $cwRow -and $c -lt $cwRow.Count) { $cwParas = $cwRow[$c] }
            if ($null -ne $hwRow -and $c -lt $hwRow.Count) { $hwParas = $hwRow[$c] }
            $split = Split-Classwork $cwParas
            $lessonContent[$n] = [pscustomobject]@{
                objective = $split.objective
                lesson    = $split.lesson
                study     = $split.study
                homework  = [string[]]$hwParas
                hasContent = (($cwParas.Count + $hwParas.Count) -gt 0)
            }
        }
    }
    else {
        $warnings += "slide${slideNo}: no CW/HW grid found ('$title')"
    }

    # --- rotation grid -> day columns ---
    $dateRow = $rotationRows[0]
    $rotRow = if ($rotationRows.Count -gt 1) { $rotationRows[1] } else { @() }
    $days = @{}
    $dayList = @()
    for ($c = 1; $c -lt $dateRow.Count; $c++) {
        $d = Parse-DateHeader -Text (Join-Cell $dateRow[$c]) -FallYear $FallSchoolYear
        if ($null -eq $d) { continue }
        $rotParas = @()
        if ($c -lt $rotRow.Count) { $rotParas = $rotRow[$c] }
        $rotation = ''
        $dayNote = ''
        if ($rotParas.Count -gt 0) {
            $rotation = $rotParas[0]
            if ($rotParas.Count -gt 1) { $dayNote = ($rotParas[1..($rotParas.Count - 1)] -join '; ') }
        }
        if ($d.weekdayMismatch) {
            $warnings += "slide${slideNo}: header says $($d.statedWeekday) for $($d.date), which is a $($d.weekday)"
        }
        $dayType = Get-DayType -Rotation $rotation -Note $notes -IsoDate $d.date
        if ($null -eq $dayType -and $rotation -ne '') {
            $warnings += "slide${slideNo}: could not map rotation '$rotation' on $($d.date) to a day type"
        }
        if ($null -ne $dayType) {
            if ($calendar.ContainsKey($d.date) -and $calendar[$d.date] -ne $dayType) {
                $warnings += "calendar conflict on $($d.date): '$($calendar[$d.date])' vs '$dayType' (slide $slideNo)"
            }
            $calendar[$d.date] = $dayType
        }
        $days[$c] = [pscustomobject]@{
            date = $d.date; weekday = $d.weekday; rotation = $rotation; dayType = $dayType; dayNote = $dayNote
        }
        $dayList += $days[$c]
    }

    $cycles += [pscustomobject]@{
        slide      = $slideNo
        title      = $title
        cycle      = $cycleNo
        isTemplate = $isTemplate
        notes      = $notes
        days       = @($dayList)
    }

    # --- period rows -> records ---
    for ($r = 2; $r -lt $rotationRows.Count; $r++) {
        $row = $rotationRows[$r]
        if ($row.Count -eq 0) { continue }
        $period = Normalize-Period (Join-Cell $row[0])
        if ($null -eq $period) { continue }

        foreach ($c in ($days.Keys | Sort-Object)) {
            if ($c -ge $row.Count) { continue }
            $cellText = Join-Cell $row[$c]
            $ln = Parse-LessonNumber $cellText
            if ($null -eq $ln) { continue }

            $day = $days[$c]
            $recWarnings = @()
            if ($ln -eq 'UNNUMBERED') {
                $recWarnings += "cell reads '$cellText' with no lesson number"
                $ln = $null
            }

            $content = $null
            if ($null -ne $ln -and $lessonContent.ContainsKey($ln)) { $content = $lessonContent[$ln] }

            $objective = ''; $lessonSteps = @(); $study = @(); $homework = @(); $hasContent = $false
            if ($null -ne $content) {
                $objective = $content.objective
                $lessonSteps = $content.lesson
                $study = $content.study
                $homework = $content.homework
                $hasContent = $content.hasContent
            }

            $records += [pscustomobject]@{
                date         = $day.date
                weekday      = $day.weekday
                rotationDay  = $day.rotation
                dayType      = $day.dayType
                dayNote      = $day.dayNote
                period       = $period
                cycle        = $cycleNo
                slide        = $slideNo
                lessonNumber = $ln
                objective    = $objective
                lesson       = [string[]]$lessonSteps
                study        = [string[]]$study
                homework     = [string[]]$homework
                notes        = $notes
                hasContent   = $hasContent
                warnings     = [string[]]$recWarnings
            }
        }
    }
}
$zip.Dispose()

return [pscustomobject]@{
    deckId     = $DeckId
    exportUrl  = $exportUrl
    deckTitle  = $deckTitle
    slideCount = $slideEntries.Count
    records    = @($records)
    cycles     = @($cycles)
    calendar   = $calendar
    warnings   = @($warnings)
}
}   # end Get-DeckSchedule

# --------------------------------------------------------------------- run ---
$allRecords = @()
$allWarnings = @()
$sources = @()
$calendar = @{}
$cycles = $null

foreach ($course in $Courses) {
    $deck = Get-DeckSchedule -DeckId $course.id -DataDir $dataDir -FallSchoolYear $FallSchoolYear

    $kept = $deck.records
    if ($course.periods) { $kept = @($deck.records | Where-Object { $course.periods -contains $_.period }) }
    foreach ($r in $kept) { Add-Member -InputObject $r -NotePropertyName 'course' -NotePropertyValue $course.label -Force }

    $covered = @($kept | Select-Object -ExpandProperty period -Unique | Sort-Object)
    if ($course.periods) {
        $absent = @($course.periods | Where-Object { $covered -notcontains $_ })
        foreach ($p in $absent) { $allWarnings += "$($course.label): no rows found for period $p" }
    }

    $allRecords += $kept
    $allWarnings += @($deck.warnings | ForEach-Object { "$($course.label): $_" })
    $sources += [pscustomobject]@{
        course = $course.label; deckId = $deck.deckId; deckTitle = $deck.deckTitle
        exportUrl = $deck.exportUrl; periods = [string[]]$course.periods
        slideCount = $deck.slideCount; recordsUsed = $kept.Count
    }

    # Every deck follows the same school rotation, so the calendars must agree.
    # A conflict means one deck's dates drifted and is worth surfacing.
    foreach ($k in $deck.calendar.Keys) {
        if ($calendar.ContainsKey($k) -and $calendar[$k] -ne $deck.calendar[$k]) {
            $allWarnings += "calendar conflict on ${k}: '$($calendar[$k])' vs '$($deck.calendar[$k])' ($($course.label))"
        }
        $calendar[$k] = $deck.calendar[$k]
    }
    if ($null -eq $cycles) { $cycles = $deck.cycles }
}

$dupes = @($allRecords | Group-Object { $_.date + '|' + $_.period } | Where-Object { $_.Count -gt 1 })
foreach ($d in $dupes) { $allWarnings += "duplicate record for $($d.Name) from $((($d.Group | Select-Object -ExpandProperty course -Unique)) -join ' + ')" }

$allRecords = $allRecords | Sort-Object date, period

$calendarOrdered = [ordered]@{}
foreach ($k in ($calendar.Keys | Sort-Object)) { $calendarOrdered[$k] = $calendar[$k] }

$payload = [pscustomobject]@{
    sources       = @($sources)
    generatedAt   = (Get-Date).ToString('o')
    schoolYear    = "$FallSchoolYear-$($FallSchoolYear + 1)"
    cycleCount    = $cycles.Count
    recordCount   = $allRecords.Count
    withContent   = (@($allRecords | Where-Object { $_.hasContent }).Count)
    warnings      = [string[]]$allWarnings
    # date -> board day-type code (1-7, X, A, B, Bn), for the board's CAL lookup.
    calendar      = $calendarOrdered
    cycles        = @($cycles)
    schedule      = @($allRecords)
}

$json = $payload | ConvertTo-Json -Depth 12
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($jsonPath, $json, $utf8NoBom)

Write-Host ""
Write-Host "Wrote $jsonPath"
foreach ($s in $sources) { Write-Host "  $($s.course): periods $($s.periods -join '/') -> $($s.recordsUsed) records" }
Write-Host "  total records=$($allRecords.Count)  with content=$($payload.withContent)  calendar days=$($calendarOrdered.Count)  warnings=$($allWarnings.Count)"
foreach ($w in $allWarnings) { Write-Host "  ! $w" }
