<#
  build-rotation.ps1

  Regenerates docs/data/rotation.json - the per-(period, date) meeting pattern
  and lesson numbering that extract-schedule.ps1 joins the sheet against.

  Supersedes tools/freeze-rotation.ps1
  --------------------------------------
  The skeleton was originally lifted out of the retired Slides decks' rotation
  grids, on the belief that the meeting pattern was recorded nowhere else. That
  was wrong: docs/index.html has always carried ROT_DAYS, a rule table saying
  exactly which periods meet on each rotation day. Deriving the pattern from
  that rule is both simpler and more correct.

  Evidence this rule is the correct source:
    * ROT_DAYS + the day-type calendar reproduce the deck-derived skeleton on
      155 of 158 school days; the 3 exceptions are the MAP testing dates, see
      below.
    * The hand-authored CAL in index.html and the generated day-type calendar
      agree on all 78 dates they share, with 0 conflicts.

  A meeting is not the same thing as a lesson
  -------------------------------------------
  The rule says who MEETS. It cannot know that a meeting is given over to
  testing, and cycle 3 has one such meeting per period. So cycle 3 has seven
  meetings and six lessons, and $NON_INSTRUCTIONAL below is what tells them
  apart.

  This file previously claimed the deck-derived skeleton was LOSSY, on the
  grounds that four cycle-3 slots were missing from it: A on 2026-09-09, G on
  09-10, D and E on 09-11. Those four are exactly the MAP testing periods. The
  decks were not losing meetings, they were declining to number them, and they
  were right to. Cycle 3's "seventh lesson" was restored on 2026-09-02 as a
  correction to that misreading; it should not have been. Do not restore it
  again without re-reading $NON_INSTRUCTIONAL first.

  The consequence of getting this wrong is subtle and easy to miss: numbering
  every meeting shifts every lesson from the testing day onward one meeting
  early, and leaves the cycle's last meeting with no content at all. The board
  renders that as a blank panel on a real teaching day.

  Inputs
  ------
    docs/index.html          ROT_DAYS and FIXED, parsed out rather than copied,
                             so this cannot drift from what the board believes.
    docs/data/rotation.json  the day-type calendar, and the date -> cycle
                             assignment. Both are kept as-is; only slots[] is
                             rebuilt. (The calendar is the one input still
                             originally deck-derived, and is cross-checked
                             against CAL above.)

  Usage:  powershell -ExecutionPolicy Bypass -File tools/build-rotation.ps1
#>
param(
    [string]$IndexPath    = (Join-Path $PSScriptRoot '..\docs\index.html'),
    [string]$RotationPath = (Join-Path $PSScriptRoot '..\docs\data\rotation.json'),
    [switch]$WhatIfOnly
)

$ErrorActionPreference = 'Stop'

$MY = @('A', 'D', 'G', 'E')
$TRACK_OF_PERIOD = @{ 'A' = 'HP'; 'D' = 'HP'; 'G' = 'HP'; 'E' = 'EP' }

# ------------------------------------------------ non-instructional meetings ---
# Meetings that happen but are not lessons. The period meets, so the rotation
# rule produces a slot for it, but nothing is taught and it must not consume a
# lesson number.
#
# MAP testing takes one period per testing day, listed by the periods pulled:
#
#   2026-09-09  DAY 1   A and B periods
#   2026-09-10  DAY 2   F and G periods
#   2026-09-11  DAY 3   D and E periods
#
# Of the four boards, that is A on 09-09, G on 09-10, and D and E on 09-11 -
# one meeting each, which is why cycle 3 has six lessons across seven meetings.
# Periods B and F are listed for completeness; no board renders them.
#
# Add future testing or assembly days here. A date with no entry is entirely
# instructional, which is the case for every other day of the year.
$NON_INSTRUCTIONAL = @{
    '2026-09-09' = @{ periods = @('A', 'B'); label = 'MAP Testing' }
    '2026-09-10' = @{ periods = @('F', 'G'); label = 'MAP Testing' }
    '2026-09-11' = @{ periods = @('D', 'E'); label = 'MAP Testing' }
}

# Returns the label if this (date, period) meeting is not a lesson, else ''.
function Get-NonInstructionalLabel {
    param([string]$Date, [string]$Period)
    if (-not $NON_INSTRUCTIONAL.ContainsKey($Date)) { return '' }
    $entry = $NON_INSTRUCTIONAL[$Date]
    if ($entry.periods -contains $Period) { return [string]$entry.label }
    return ''
}

# ------------------------------------------------- parse the board's tables ---
$html = Get-Content $IndexPath -Raw -Encoding UTF8

function Get-JsBlock {
    param([string]$Text, [string]$Declaration)
    $i = $Text.IndexOf($Declaration)
    if ($i -lt 0) { throw "Could not find '$Declaration' in $IndexPath" }
    $j = $Text.IndexOf('};', $i)
    if ($j -lt 0) { throw "Unterminated '$Declaration' in $IndexPath" }
    return $Text.Substring($i, $j - $i)
}

# ROT_DAYS: '<dayType>': ['A','B',...]
$rotBlock = Get-JsBlock $html 'const ROT_DAYS = {'
$ROT_DAYS = @{}
foreach ($m in [regex]::Matches($rotBlock, "'([^']+)'\s*:\s*\[([^\]]*)\]")) {
    $periods = @()
    foreach ($p in [regex]::Matches($m.Groups[2].Value, "'([^']+)'")) { $periods += $p.Groups[1].Value }
    $ROT_DAYS[$m.Groups[1].Value] = $periods
}
if ($ROT_DAYS.Count -eq 0) { throw 'Parsed no entries from ROT_DAYS' }

# FIXED: <dayType>: { <period>:'time', ... } - here the periods are the keys.
$fixedBlock = Get-JsBlock $html 'const FIXED = {'
$FIXED = @{}
foreach ($line in ($fixedBlock -split "`r?`n")) {
    $m = [regex]::Match($line, '^\s*(\w+)\s*:\s*\{(.+)\}')
    if (-not $m.Success) { continue }
    $periods = @()
    foreach ($p in [regex]::Matches($m.Groups[2].Value, "(\w+)\s*:\s*'")) { $periods += $p.Groups[1].Value }
    if ($periods.Count) { $FIXED[$m.Groups[1].Value] = $periods }
}
if ($FIXED.Count -eq 0) { throw 'Parsed no entries from FIXED' }

Write-Host ("Parsed from index.html: ROT_DAYS {0} day types, FIXED {1} day types" -f $ROT_DAYS.Count, $FIXED.Count)

function Get-MeetingPeriods {
    param([string]$DayType)
    $set = $null
    if ($ROT_DAYS.ContainsKey($DayType)) { $set = $ROT_DAYS[$DayType] }
    elseif ($FIXED.ContainsKey($DayType)) { $set = $FIXED[$DayType] }
    if ($null -eq $set) { return @() }
    return @($MY | Where-Object { $set -contains $_ })
}

# --------------------------------------------- calendar and cycle boundaries ---
$existing = Get-Content $RotationPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $existing.calendar) { throw "No calendar in $RotationPath" }
if (-not $existing.slots)    { throw "No slots[] in $RotationPath (needed for the date -> cycle map)" }

$dateCycle = @{}
foreach ($s in $existing.slots) {
    $k = $s.date
    $v = [pscustomobject]@{ semester = [int]$s.semester; cycle = [int]$s.cycle }
    if ($dateCycle.ContainsKey($k)) {
        if ($dateCycle[$k].semester -ne $v.semester -or $dateCycle[$k].cycle -ne $v.cycle) {
            throw "Date $k is assigned to more than one cycle"
        }
    }
    $dateCycle[$k] = $v
}

$calDates = @($existing.calendar.PSObject.Properties.Name | Sort-Object)

# ------------------------------------------------------------ rebuild slots ---
# Lesson numbers are the Nth meeting of that period within its cycle, so the
# dates must be walked in order per (cycle, period).
$byCyclePeriod = @{}
$unassigned = @()
foreach ($date in $calDates) {
    $dayType = [string]$existing.calendar.$date
    if (-not $dateCycle.ContainsKey($date)) { $unassigned += $date; continue }
    $cy = $dateCycle[$date]
    foreach ($period in (Get-MeetingPeriods $dayType)) {
        $k = '{0}|{1}|{2}' -f $cy.semester, $cy.cycle, $period
        if (-not $byCyclePeriod.ContainsKey($k)) { $byCyclePeriod[$k] = @() }
        $byCyclePeriod[$k] += [pscustomobject]@{ date = $date; dayType = $dayType }
    }
}
if ($unassigned.Count) {
    Write-Warning ("{0} calendar date(s) belong to no cycle and were skipped: {1}" -f
        $unassigned.Count, ($unassigned -join ', '))
}

# Carry the human-facing rotation label and any day note across from the old
# skeleton, keyed on (date, period); a newly added slot inherits its date's.
$oldByDatePeriod = @{}
$oldByDate = @{}
foreach ($s in $existing.slots) {
    $oldByDatePeriod['{0}|{1}' -f $s.date, $s.period] = $s
    if (-not $oldByDate.ContainsKey($s.date)) { $oldByDate[$s.date] = $s }
}

$slots = New-Object System.Collections.ArrayList
$added = @()
$nonInstructional = @()
foreach ($k in $byCyclePeriod.Keys) {
    $parts = $k.Split('|')
    $semester = [int]$parts[0]; $cycle = [int]$parts[1]; $period = $parts[2]
    $n = 0
    foreach ($d in ($byCyclePeriod[$k] | Sort-Object date)) {
        # A testing meeting still gets a slot - the class does meet, and the
        # board should not silently skip the day - but it takes no lesson
        # number, so the lessons after it keep their real dates.
        $label = Get-NonInstructionalLabel -Date $d.date -Period $period
        $lessonNumber = $null
        if ($label) {
            $nonInstructional += ('{0} period {1}  {2}' -f $d.date, $period, $label)
        } else {
            $n++
            $lessonNumber = $n
        }

        $ref = $oldByDatePeriod['{0}|{1}' -f $d.date, $period]
        if (-not $ref) {
            $ref = $oldByDate[$d.date]
            $added += ('{0} period {1}' -f $d.date, $period)
        }
        $weekday = ''; $rotationDay = ''; $dayNote = ''
        if ($ref) { $weekday = $ref.weekday; $rotationDay = $ref.rotationDay; $dayNote = $ref.dayNote }
        # dayNote is per-slot, so it is the natural home for the label. Keep any
        # carried note and put the label in front of it.
        if ($label) {
            if ($dayNote -and $dayNote -notlike "*$label*") { $dayNote = '{0} - {1}' -f $label, $dayNote }
            elseif (-not $dayNote) { $dayNote = $label }
        }

        [void]$slots.Add([pscustomobject]@{
            date         = $d.date
            weekday      = $weekday
            period       = $period
            track        = $TRACK_OF_PERIOD[$period]
            semester     = $semester
            cycle        = $cycle
            lessonNumber = $lessonNumber
            rotationDay  = $rotationDay
            dayType      = $d.dayType
            dayNote      = $dayNote
        })
    }
}
$sorted = $slots | Sort-Object date, period

# ------------------------------------------------------------- validation ---
$dupes = $sorted | Group-Object { $_.period + '|' + $_.date } | Where-Object { $_.Count -gt 1 }
if ($dupes.Count) { throw ("Duplicate (period, date): {0}" -f (($dupes | ForEach-Object { $_.Name }) -join ', ')) }

# Lesson numbers must still run 1..N with no gaps, counting only the meetings
# that are lessons. A testing slot carries $null and is excluded here; if one
# ever leaked a number in, this is what would catch it.
foreach ($g in ($sorted | Group-Object { '{0}|{1}|{2}' -f $_.semester, $_.cycle, $_.period })) {
    $nums = @($g.Group | Where-Object { $null -ne $_.lessonNumber } |
              ForEach-Object { [int]$_.lessonNumber } | Sort-Object)
    if ($nums.Count -eq 0) { throw ("No lessons at all for {0}" -f $g.Name) }
    if (Compare-Object $nums (1..$nums.Count)) {
        throw ("Lesson numbers for {0} are not 1..{1}: {2}" -f $g.Name, $nums.Count, ($nums -join ','))
    }
}

$lessonSlots = @($sorted | Where-Object { $null -ne $_.lessonNumber })

Write-Host ""
Write-Host ("slots: {0} (was {1}); newly added: {2}" -f $sorted.Count, @($existing.slots).Count, $added.Count)
foreach ($a in ($added | Sort-Object)) { Write-Host ("  + {0}" -f $a) }
Write-Host ""
Write-Host ("non-instructional meetings (slot, but no lesson number): {0}" -f $nonInstructional.Count)
foreach ($t in ($nonInstructional | Sort-Object)) { Write-Host ("  = {0}" -f $t) }
Write-Host ""
Write-Host ("lesson slots: {0}  (this is the number Master Data should have rows for)" -f $lessonSlots.Count)
Write-Host ""
Write-Host "max lesson number per (semester|cycle|track):"
foreach ($g in ($lessonSlots | Group-Object { '{0}|{1}|{2}' -f $_.semester, $_.cycle, $_.track } | Sort-Object Name)) {
    $max = (($g.Group | ForEach-Object { [int]$_.lessonNumber }) | Measure-Object -Maximum).Maximum
    Write-Host ("  {0} -> {1}" -f $g.Name, $max)
}

if ($WhatIfOnly) { Write-Host ""; Write-Host 'WhatIfOnly - nothing written.'; return }

$out = [ordered]@{
    note        = 'Per-period meeting dates and lesson numbering for the school year. Generated by tools/build-rotation.ps1 from the ROT_DAYS rule table in docs/index.html plus the day-type calendar below. Do not hand-edit; re-run the generator.'
    generatedAt = (Get-Date).ToUniversalTime().ToString('yyyy-MM-ddTHH:mm:ssZ')
    schoolYear  = $existing.schoolYear
    derivedFrom = 'docs/index.html ROT_DAYS + FIXED, and the day-type calendar in this file'
    secondSemesterStart = $existing.secondSemesterStart
    slotCount   = $sorted.Count
    calendar    = $existing.calendar
    slots       = $sorted
}
$json = $out | ConvertTo-Json -Depth 6
$json = $json -replace "`r`n", "`n"

# Same rule as the extractor: when only the clock moved, keep the old bytes, so
# re-running the generator to check it does not produce a phantom diff.
$stampOnly = $false
if (Test-Path $RotationPath) {
    $stampRx = '"generatedAt":\s*"[^"]*"'
    $old = [System.IO.File]::ReadAllText($RotationPath)
    if ([regex]::Replace($json, $stampRx, '""') -eq [regex]::Replace($old, $stampRx, '""')) {
        $json = $old
        $stampOnly = $true
    }
}

$noBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($RotationPath, $json, $noBom)
Write-Host ""
if ($stampOnly) { Write-Host ("unchanged - {0} left byte-identical" -f (Resolve-Path $RotationPath)) }
else { Write-Host ("wrote {0}" -f (Resolve-Path $RotationPath)) }
