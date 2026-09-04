<#
  Generates the cell-level fix for the "Lesson Planner" tab of the planning
  sheet: restores the line breaks and the Lit. Focus / Lesson split that were
  lost when the sheet was first built from the Slides decks.

  Correct content comes from docs/data/schedule.json, which parsed the decks
  properly (objective / lesson[] / study[] / homework[]).

  Scope: Semester 1, cycles 1-2 only. Those are the only pre-filled cycles.
    cycle 1 -> Lit. Focus was empty and Lesson held everything, so all four
               content fields are rewritten.
    cycle 2 -> field separation is already correct (and the human-curated
               Lit. Focus reads better than the deck parse), so only the
               multi-line fields are rewritten.

  IMPORTANT - this needs a DECK-derived schedule.json
  ---------------------------------------------------
  It reads the correct, properly-split content out of schedule.json. Today's
  schedule.json is built from the sheet, whose cycle 1-2 content is the very
  thing being repaired, so running this against the current file produces
  nothing useful. To regenerate, first recover a deck-era schedule.json from
  git history (the last "chore: refresh schedule from slides deck" commit) and
  pass it via -SchedulePath.

  The generated repair is folded into tools/apply-sheet-fixes.gs, so this
  script is kept only for provenance and re-generation.
#>
param(
    [string]$SchedulePath = (Join-Path $PSScriptRoot '..\docs\data\schedule.json'),
    [string]$OutDir = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'
$LF = [string][char]10

$d = Get-Content $SchedulePath -Raw -Encoding UTF8 | ConvertFrom-Json

# The deck leaked its own row label into some objective cells, and left
# trailing colons where the slide used "Digital Setup:" as a heading.
function Clean-Focus {
    param([string]$s)
    if ($null -eq $s) { return '' }
    $s = $s -replace '^\s*Lit\.\s*Focus:\s*', ''
    $s = $s.Trim()
    $s = $s.TrimEnd(':')
    return $s.Trim()
}

# Lesson Planner layout: one 18-row block per cycle starting at row 1.
# Within a block, offset from the block's first row:
#   +2 HP First 5   +3 Bring  +4 Lit. Focus  +5 Lesson  +6 Study  +7 Homework  +8 Due
#   +9 EP First 5  +10 Bring +11 Lit. Focus +12 Lesson +13 Study +14 Homework +15 Due
$ROW_LITFOCUS = 4
$ROW_LESSON   = 5
$ROW_STUDY    = 6
$ROW_HOMEWORK = 7
$EP_OFFSET    = 7
$BLOCK_STRIDE = 18

# Lesson 1 sits in column C.
function Col-For {
    param([int]$LessonNumber)
    return [string][char]([int][char]'C' + $LessonNumber - 1)
}

# Cells where the deck parse is WORSE than what the teacher already has, so
# rewriting them would be a regression. Left alone deliberately.
#   H24 - deck paragraphs came out reordered: the story title landed on its own
#         line ahead of "Continue -Read". The sheet's curated single line
#         ("Continue "The Fan Club"-Discuss") already reads correctly.
$EXCLUDE = @('H24')

$cells = New-Object System.Collections.ArrayList

foreach ($cy in 1, 2) {
    $blockStart = 1 + $BLOCK_STRIDE * ($cy - 1)
    foreach ($track in 'HP', 'EP') {
        if ($track -eq 'HP') { $periods = @('A', 'D', 'G'); $off = 0 }
        else                 { $periods = @('E');           $off = $EP_OFFSET }

        # Semester 1 only: `cycle` repeats across semesters, so filter by date.
        $rows = $d.schedule | Where-Object {
            $_.cycle -eq $cy -and $periods -contains $_.period -and $_.date -lt '2027-01-01'
        }

        foreach ($g in ($rows | Group-Object lessonNumber | Sort-Object { [int]$_.Name })) {
            $r = $g.Group[0]
            $n = [int]$g.Name
            $col = Col-For $n

            $fields = New-Object System.Collections.ArrayList
            if ($cy -eq 1) {
                [void]$fields.Add(@{ name = 'Lit. Focus'; row = $ROW_LITFOCUS; value = (Clean-Focus $r.objective) })
            }
            [void]$fields.Add(@{ name = 'Lesson';   row = $ROW_LESSON;   value = (($r.lesson)   -join $LF) })
            [void]$fields.Add(@{ name = 'Study';    row = $ROW_STUDY;    value = (($r.study)    -join $LF) })
            [void]$fields.Add(@{ name = 'Homework'; row = $ROW_HOMEWORK; value = (($r.homework) -join $LF) })

            foreach ($f in $fields) {
                if ([string]::IsNullOrWhiteSpace($f.value)) { continue }
                $a1 = $col + [string]($blockStart + $f.row + $off)
                if ($EXCLUDE -contains $a1) {
                    Write-Host ("skipping {0} (excluded: deck parse is worse)" -f $a1)
                    continue
                }
                [void]$cells.Add([pscustomobject]@{
                    cell   = $a1
                    field  = $f.name
                    cycle  = $cy
                    track  = $track
                    lesson = $n
                    lines  = ($f.value -split $LF).Count
                    value  = $f.value
                })
            }
        }
    }
}

$multi = ($cells | Where-Object { $_.lines -gt 1 }).Count
Write-Host ("cells to write : {0}" -f $cells.Count)
Write-Host ("multi-line     : {0}" -f $multi)
Write-Host ""
$cells | Select-Object cell, cycle, track, lesson, field, lines | Format-Table -AutoSize

$jsonPath = Join-Path $OutDir 'fix-cells.json'
$cells | ConvertTo-Json -Depth 4 | Set-Content -Path $jsonPath -Encoding utf8
Write-Host ("wrote {0}" -f $jsonPath)

# ------------------------------------------------------- Apps Script emitter ---
# A one-shot script the teacher runs from the sheet itself (Extensions > Apps
# Script). Writes only the cells listed above; every other cell, all formatting
# and the Master Data formulas are untouched.
$sb = New-Object System.Text.StringBuilder
[void]$sb.AppendLine('/**')
[void]$sb.AppendLine(' * One-shot repair for the "Lesson Planner" tab.')
[void]$sb.AppendLine(' *')
[void]$sb.AppendLine(' * Restores the line breaks (and, for cycle 1, the Lit. Focus / Lesson split)')
[void]$sb.AppendLine(' * that were flattened when this sheet was first built from the Slides decks.')
[void]$sb.AppendLine(' * Touches only the cells listed in FIXES. Master Data is formula-driven and')
[void]$sb.AppendLine(' * will update itself.')
[void]$sb.AppendLine(' *')
[void]$sb.AppendLine(' * To run: Extensions > Apps Script, paste this file, Run > repairLessonPlanner.')
[void]$sb.AppendLine(' */')
[void]$sb.AppendLine('')
[void]$sb.AppendLine("var SHEET_NAME = 'Lesson Planner';")
[void]$sb.AppendLine('')
[void]$sb.AppendLine('var FIXES = [')
foreach ($c in $cells) {
    $esc = $c.value.Replace('\', '\\').Replace("'", "\'").Replace([string][char]10, '\n')
    $note = "cycle {0} {1} L{2} {3}" -f $c.cycle, $c.track, $c.lesson, $c.field
    [void]$sb.AppendLine(("  {{ a1: '{0}', v: '{1}' }},  // {2}" -f $c.cell, $esc, $note))
}
[void]$sb.AppendLine('];')
[void]$sb.AppendLine('')
[void]$sb.AppendLine('function repairLessonPlanner() {')
[void]$sb.AppendLine('  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);')
[void]$sb.AppendLine('  if (!sh) { throw new Error(''No sheet named '' + SHEET_NAME); }')
[void]$sb.AppendLine('  for (var i = 0; i < FIXES.length; i++) {')
[void]$sb.AppendLine('    sh.getRange(FIXES[i].a1).setValue(FIXES[i].v).setWrap(true);')
[void]$sb.AppendLine('  }')
[void]$sb.AppendLine('  SpreadsheetApp.getActive().toast(FIXES.length + '' cells repaired'', ''Lesson Planner'', 5);')
[void]$sb.AppendLine('}')

$gsPath = Join-Path $OutDir 'repair-lesson-planner.gs'
# PS 5.1's -Encoding utf8 emits a BOM; the Apps Script editor shows it as stray
# glyphs on line 1, so write UTF-8 without one.
$noBom = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllText($gsPath, $sb.ToString(), $noBom)
Write-Host ("wrote {0}" -f $gsPath)
