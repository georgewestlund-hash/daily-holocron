# Daily Holocron — Classroom Planner

A classroom display board for 8th-grade Literature. `docs/` is a static site
(GitHub Pages) showing one panel per class period for the current date, with
lesson content pulled from a generated JSON file. Styling is Star Wars-flavoured
(starfield, laser dividers, "Holocron").

## Data flow

The planning sheet is the database. It has two consumers, and they are
independent — neither knows about the other, and either can be run without the
other.

```
                                            ┌─[extract-schedule.ps1]─▶ docs/data/schedule.json ─[fetch]─▶ docs/index.html
planning Google Sheet ("Master Data" tab) ──┤                                (committed)                     the board
                                            └─[tools/sync-sheet-to-slides.gs]─▶ the two Classwork/Homework decks
docs/data/rotation.json (frozen skeleton) ──┘                                    (Google Slides, in Drive)
```

The rotation skeleton feeds only the board. The decks carry their own rotation
grids and their own lesson numbering, read at run time — see
*The slide decks*, below.

`.github/workflows/refresh.yml` runs the extractor every 6 hours, commits
`schedule.json` **only if it changed**, and deploys `docs/` to Pages.

Pages publishing is **opt-in**: the deploy steps are gated on the repo variable
`ENABLE_PAGES == 'true'`. A private repo on a free plan cannot serve Pages, so
leaving it on would fail every run.

## The two inputs, and why there are two

**The sheet has no dates.** `Master Data` is keyed on
`(Semester, Cycle, Lesson #, Track)`. The board needs one record per
`(period, date)`. The bridge is the per-period meeting pattern, and it is not
in the sheet:

- Periods do not all meet every rotation day, and they skip *different* days.
  In cycle 1, period A meets on rotation days 1, 2, 3, 5 and 6 but not 4 or 7.
- The rotation *label's* parenthetical is not the meeting list either. On
  2026-08-17 the label is `DAY 2 (F-C)`, yet A and G are the periods that meet.

**The pattern is a rule, and the rule lives in the board.**
`docs/index.html` line ~319 has `ROT_DAYS`, mapping each rotation day to the
five periods that meet, plus `FIXED` for the X / A-block / B-block days:

```
'1': A B C D E      '5': G A B C D
'2': F G A B C      '6': E F G A B
'3': D E F G A      '7': C D E F G
'4': B C D E F      X: all seven   A: A B C D   B/Bn: E F G
```

`tools/build-rotation.ps1` parses those tables straight out of `index.html`
(never copied, so they cannot drift) and crosses them with the day-type
calendar to produce `docs/data/rotation.json` — 440 slots, lesson number =
the Nth meeting of that period in its cycle.

### Do not go back to the decks for this

The skeleton was originally *frozen* from the retired Slides decks' rotation
grids, on the belief that the pattern was recorded nowhere else. That was
wrong: `ROT_DAYS` had it all along, and deriving the pattern from the rule is
both simpler and more correct.

Two independent checks:

| Check | Result |
|---|---|
| `ROT_DAYS` vs the deck-derived skeleton | agree on **155 of 158** school days; the 3 exceptions are the MAP testing dates — see *A meeting is not a lesson* |
| Hand-authored `CAL` vs the generated day-type calendar | **0 conflicts** across the 78 dates they share |

```
            c1 c2 c3 c4 c5 c6 c7 c8 c9
lessons  S1  6  6  6  7  7  5  6  5  6      per track, 54
lessons  S2  6  6  6  6  6  6  6  7  6      per track, 55
```

### A meeting is not a lesson

**The rule says who *meets*. It cannot know that a meeting is given over to
testing.** MAP testing pulls one period per testing day:

```
2026-09-09  DAY 1   A and B periods
2026-09-10  DAY 2   F and G periods
2026-09-11  DAY 3   D and E periods
```

Of the four boards that is A on 09-09, G on 09-10, and D and E on 09-11 — one
meeting each. So **S1 cycle 3 has seven meetings and six lessons**, and
`$NON_INSTRUCTIONAL` in `build-rotation.ps1` is what tells them apart. Such a
meeting still gets a slot (the class does meet, and the board should not skip
the day) with `lessonNumber = null` and `dayNote = 'MAP Testing'`, and it does
not consume a lesson number.

Getting this wrong is subtle. Numbering every meeting shifts every lesson from
the testing day onward one meeting early and leaves the cycle's last meeting
empty — a blank panel on a real teaching day:

```
period A     numbered by meeting        corrected
9/08  Tue    Lesson 1                   Lesson 1
9/09  Wed    Lesson 2                   MAP Testing
9/10  Thu    Lesson 3                   Lesson 2
9/11  Fri    Lesson 4                   Lesson 3
9/15  Tue    Lesson 5                   Lesson 4
9/16  Wed    Lesson 6                   Lesson 5
9/18  Fri    Lesson 7  → blank          Lesson 6
```

**This corrects a reversal that went the wrong way.** On 2026-09-02 the
deck-derived skeleton was judged *lossy* because four cycle-3 slots were absent
from it — A on 09-09, G on 09-10, D and E on 09-11. Those four are exactly the
MAP testing periods. The decks were not losing meetings, they were declining to
number them, and they were right. Cycle 3's "seventh lesson" was restored as a
correction to that misreading and should not have been. The teacher settled it
on 2026-09-04: *"Because of MAP testing, cycle 3 only has 6 lessons."*

Consequently the `Cycle Calendar` `# Lessons` check is no longer a clean 18 of
18 — `E4` still says 7 where the rebuilt count is now 6. See
[To finish the switch](#to-finish-the-switch).

The day-type `calendar` inside `rotation.json` is the one part still originally
deck-derived; `build-rotation.ps1` carries it forward untouched. It is
cross-checked against `CAL` (row 3 above), so it is trustworthy — but if the
school calendar ever changes, that map is what needs updating, and everything
else regenerates from it.

### 1. The planning Google Sheet

Link-shared; `export?format=xlsx` is fetched anonymously, the same trust model
the decks used.

**The sheet id is deliberately not recorded anywhere in this repo** — see
[Keeping the sheet unlisted](#keeping-the-sheet-unlisted). It is read from
`$env:HOLOCRON_SHEET_ID` (the `SHEET_ID` Actions variable in CI) or from a
gitignored `sheet-id.local` file beside the extractor. Do not paste it into
source, docs, or commit messages.

Tabs: `Lesson Planner` (human-facing, one 18-row block per cycle, merged cells),
`Cycle Calendar` (cycle → date range + lesson count), `Master Data`, `Read Me`.

**Read `Master Data`, never `Lesson Planner`.** The sheet's own Read Me
designates it as the machine-readable tab: flat, formula-driven, one row per
lesson, no merged cells. Column order is the contract:

```
A Semester | B Cycle | C Lesson # | D Track | E First 5 | F Bring
G Lit. Focus | H Lesson Plan | I Study | J Homework | K Due
```

**Gotcha: C and H were both originally headed `Lesson`** — C is the lesson
*number*, H is the lesson *content*. Address columns by letter, never by header
name; a name-keyed lookup silently collides and makes every row look like it has
content. That bug is why the header row is being renamed to `Lesson #` /
`Lesson Plan`.

The extractor validates all eleven headers against `$COLS` and throws rather
than guessing if the layout moves. `$COLS` maps each letter to a *list* of
acceptable names, so C accepts `Lesson #` or `Lesson` and H accepts
`Lesson Plan` or `Lesson`. The sheet can therefore be renamed at any time,
in either order, without a matching code change — while a real change (a typo,
two columns swapped) is still rejected.

`Track` is `HP` (Honors — periods A, D, G, planned in lockstep) or `EP`
(ECP — period E, paced separately).

Blank formula cells export as `t="str"` with an empty `<v/>`, so they read as
empty strings and fall through as blank records — which the board renders as an
empty panel. That is intended: 220 rows exist, only the filled ones count. (218
of them correspond to a real lesson; the two extras are S1 cycle 3 lesson 7 —
see [The health check worth re-running](#the-health-check-worth-re-running).)

### 2. `docs/data/rotation.json`

Committed. One slot per `(period, date)` carrying `semester`, `cycle`,
`lessonNumber`, `track`, `rotationDay`, `dayType`, `dayNote`, plus the
`calendar` map. Do not hand-edit.

440 slots, of which **436 are lessons and 4 are MAP testing** — those four
carry `lessonNumber = null`. Collapsed to join keys that is **218**
`(semester, cycle, lessonNumber, track)` combinations, because HP's three
periods share one key per lesson.

**`cycle` is not unique** — semester 2 restarts at cycle 1. Every join must key
on `(semester, cycle, lessonNumber, track)`, or the two Cycle 1s collide.

**`lessonNumber` can be null.** A null means the class meets but no lesson is
taught. The extractor's key then matches nothing, the record publishes with
`hasContent: false`, and the board renders a blank panel — correct, because
there is nothing to teach. Anything joining on `lessonNumber` must tolerate it.

## The board (`docs/index.html`)

Single file, ~80 KB, one big class component with inline styles. Reads
`data/schedule.json` once on load (`loadDeck`, ~line 550). Internal identifiers
still say "deck" for historical reasons; the source is the sheet.

- `const MY = ['A','D','G','E']` — the four period boards rendered.
- `const CAL = {...}` (~line 334) — hand-authored day-type calendar carrying
  holiday/exam labels the sheet cannot express. **CAL wins**; the generated
  `calendar` only fills gaps (notably all of second semester).
- Fields are `contenteditable` elements keyed by `data-key`. Those marked
  `data-daily="1"` persist per-date (storage key suffixed with the date).
- `applyDeck()` (~line 588) fills the seven sheet-backed keys. Anything the
  teacher has typed for that date wins unless `force` is set, and deck text is
  deliberately never written to storage — a cached copy would be
  indistinguishable from a real edit on the next load.
- `FIT_KEYS` auto-shrink type to fit their panel rather than scrolling. `Bring`
  and `Due` are deliberately excluded (single-line fields).

### Field keys

| Board key   | Sheet column | Type            |
|-------------|--------------|-----------------|
| `first5`    | First 5      | string          |
| `materials` | Bring        | string          |
| `objective` | Lit. Focus   | string          |
| `lesson`    | Lesson (H)   | array of lines  |
| `study`     | Study        | array of lines  |
| `homework`  | Homework     | array of lines  |
| `due`       | Due          | string          |

All seven are sheet-backed. The three array fields are joined with newlines and
assigned via `el.innerText`, so **line structure is meaningful** — a
space-joined blob renders as one run-on line. Multi-line sheet cells are split
on newlines; blank lines are dropped.

## The slide decks (`tools/sync-sheet-to-slides.gs`)

The sheet's second consumer. Two Google Slides decks — **8 Honors Literature**
(periods A, D, G; track `HP`) and **8 ECP Literature** (period E; track `EP`) —
carry one or two slides per cycle, each with a CW/HW table. The script pushes
four sheet columns into those cells:

| Sheet column | Deck cell | Shape |
|---|---|---|
| Lit. Focus  | CW | behind a `Lit. Focus:` label |
| Lesson Plan | CW | behind a `Lesson:` label |
| Study       | CW | behind a `Study:` label |
| Homework    | HW | one sheet line per item, blank line between |

The three CW sections are separated by a blank line, and a field with no
content contributes nothing — no orphan label, no stray blank line.

**These are the two retired source decks.** They fed the board before the sheet
migration; the flow is now the other way round. Their ids are already public
(committed history, and deployed `schedule.json` versions), so `CONFIG.TRACKS`
names them openly. The sheet id is a different matter: the script is **bound to
the spreadsheet**, so `SpreadsheetApp.getActive()` reads Master Data and no id
appears in the file. Keep `CONFIG.SPREADSHEET_ID` empty — see
[Keeping the sheet unlisted](#keeping-the-sheet-unlisted).

### How a slide is matched to sheet rows

The join key is `(semester, cycle, lesson #, track)` — the same key the
extractor uses, and for the same reason: **`cycle` is not unique**, since
semester 2 restarts at cycle 1.

- **track** — from the deck id, via `CONFIG.TRACKS`.
- **cycle** — `Cycle 3` anywhere in the slide's title box.
- **semester** — worked out from the title's date range, cross-checked against
  an `S1`/`S2` in the label when there is one, and a disagreement is reported.
  Both are needed: the Honors deck writes `– S1 - Cycle 3`, the ECP deck writes
  `- S1 Cycle 3`, and the ECP deck's first two slides give no semester at all.
- **lesson #** — read off the CW/HW table's own `Lesson N` header cells.

**Never infer the lesson number from column position.** Cycles 3, 4, 5 and 8
are split across two slides, and the second slide's first data column is Lesson
4 or Lesson 5. Position-based mapping silently writes every one of those cycles
into the wrong columns, and the text looks plausible either way.

### It only fills empty cells

`CONFIG.ONLY_FILL_EMPTY_CELLS` is `true`. Anything already typed into a deck
cell is kept; the run still reports what the sheet would have put there, so a
divergence is visible without being acted on. The sync therefore *adds* the
lessons a deck is missing and changes nothing already written.

`CONFIG.BLANK_SHEET_CLEARS_SLIDE` is `false` for the same reason from the other
direction: an empty sheet cell never wipes a slide cell.

Both matter because Master Data is filled for **S1 cycles 1–2 only**, while the
decks still hold hand-written text the sheet has never had — all of S1 cycle 3,
and `No HW` / `No Homework Weekend!` notes scattered further out.

**The trade-off, and it is a real one:** a filled cell is now beyond the sheet's
reach. Edit a lesson in Master Data and the board updates on the next refresh
while the deck does not, so the two can drift apart silently. Setting
`ONLY_FILL_EMPTY_CELLS` to `false` makes the sheet win wherever both sides have
text, which is what genuine single-source authority needs. Preview reports the
divergence either way, and that report is the thing to read before deciding.

### When the sheet has a lesson the deck has no column for

The run reports it, once per cycle per deck, and will not invent a column.

The check is deck-wide, computed in a first pass over all slides, precisely
because a lesson missing from one slide is normally just on its sibling — cycles
3, 4, 5 and 8 are split across two.

**Its one firing so far was a false alarm worth remembering.** S1 cycle 3 has
six lesson columns in both decks, and while the sheet briefly carried a seventh
row the run flagged it as having nowhere to go. The decks were right: cycle 3
has six lessons, because MAP testing takes one of its seven meetings. So a
report from this check is a question — *does that lesson really exist?* — and
not automatically a deck defect. See *A meeting is not a lesson*.

### Running it

`preview` writes nothing and reports every cell it would change with old and
new text. `listSlides` prints what the script sees on each slide — the parsed
title, the table it picked, and the column→lesson map — and is the thing to run
when a slide is skipped for a reason that is not obvious. A live run saves a
dated backup copy of each deck to Drive first, and Slides' own version history
is a second net. Re-running writes nothing: composed text compares equal to
text already synced.

Character styling is captured before each write and reapplied, so a synced cell
keeps its font, size and colour. `tools/format-slide-decks.gs` is the separate
tool for table geometry and cell shading; the two do not overlap.

## `schedule.json` contract

The interface between extractor and board. Keep the shape stable.

```
sources[]     two entries: the sheet (id, tab, row counts) and the skeleton
generatedAt   ISO timestamp — see the churn note below
schoolYear    e.g. "2026-2027"
recordCount   total records (= rotation slot count)
withContent   how many matched a filled Master Data row
matchedSlots  same, from the join's side — a sanity pair
warnings[]    non-fatal complaints, surfaced for the teacher
calendar      { "YYYY-MM-DD": dayType }   dayType = "1".."7" | "A" | "B" | "X"
schedule[]    one record per (period, date) — the board's index
              { date, weekday, rotationDay, dayType, dayNote, period, track,
                semester, cycle, lessonNumber, first5, materials, objective,
                lesson[], study[], homework[], due, hasContent, course }
```

The board indexes `schedule[]` by `[period][date]`; anything absent renders
blank. It reads only `schedule[]`, `calendar`, `generatedAt` and the seven
content fields — the rest is provenance and diagnostics.

Dropped in the migration (nothing read them): the top-level `cycles[]` array and
`cycleCount`, and the per-record `slide`, `notes` and `warnings` fields.

## Publication horizon — built, currently OFF

`-PublishLeadDays` withholds a cycle's lesson content until today is within N
days of its first meeting, so unfinished future planning stays out of a
publicly served file. Withheld slots publish blank — identical to an unplanned
lesson, which the board already renders as an empty panel.

**Default is `0` (off), deliberately.** The board's URL is not shared with
students or families, so there is no audience to withhold drafts from, and
turning it on costs the ability to page the board forward when prepping — a
real loss for no gain. Change the default to `2` (or `7`) if the address is
ever handed out. The mechanism is finished and tested; only the default is off.

- The current cycle is always open, so same-day edits still land on the next
  6-hourly run. The horizon changes *what* is published, never how often.
- Cycle starts come from `rotation.json` (earliest slot date per
  `semester|cycle`), **not** from the sheet's `Cycle Calendar` — which is the
  tab that had the cycle-3 miscount.
- "Today" is the runner's **UTC** date. Eastern is UTC-4/-5, so a cycle opens a
  few hours early in local terms; immaterial against a whole-day lead, and it
  avoids the Windows/Linux timezone-id split that would make local and CI runs
  disagree.
- The run report lists every withheld cycle with its open date, and
  `withheldSlots` / `publishLeadDays` are recorded in the payload.
- `-Today <yyyy-MM-dd>` is a testing hook for exercising the boundary. Verified
  inclusive at `start - lead`: for cycle 1|2 (starts 8/25) content is absent on
  8/22, present on 8/23, and byte-identical on 8/24.
- `-PublishLeadDays 0` disables the horizon.

**Trade-off:** the board cannot be paged forward past the horizon to preview.
Raising the lead is the lever if that becomes annoying.

Subtlety worth preserving: a withheld cycle's Master Data rows are still marked
*used* for phantom-row detection. Skipping that would raise a "no matching
rotation slot" warning for every withheld row and bury the real mistakes —
which, with warnings now failing the build, would mean a red run every day.

## Keeping the sheet unlisted

This repo is **public** and `docs/` is served on GitHub Pages, so four separate
channels will leak anything careless. The planning sheet holds the whole year's
teacher-side planning and is link-shared, so its id is kept out of all four:

| Channel | Why it leaks | Guard |
|---|---|---|
| Source and docs | Public repo, and public git history | Id read from `$env:HOLOCRON_SHEET_ID` or gitignored `sheet-id.local`; never written down |
| `schedule.json` | Published at `/data/schedule.json` | `sources[]` records the tab and row counts only — no id, no export URL |
| Actions logs | Public on a public repo | Only `$sheetTag` (last 4 chars) is ever printed, never `$exportUrl` |
| The Pages artifact | `upload-pages-artifact` uploads the whole `docs/` tree **as it stands on the runner — gitignored files included** | Downloads go to `.cache/`, outside `docs/`; the extractor also deletes any stray `deck-*.pptx` / `sheet-*.xlsx` / `*.tmp` under `docs/data` before writing |

That last one is the non-obvious one, and it bit this project for real: because
earlier versions downloaded the decks into `docs/data/`, both were served in
full from the live site (`/data/deck-<id>.pptx`, HTTP 200, ~160 KB each) despite
being gitignored. Gitignore does not protect the published tree.

**This is obscurity, not access control.** Anyone holding the link can still
open the sheet; what is closed off is the realistic discovery path of reading
the id out of the public repo or the published JSON. Real protection means the
sheet stops being link-shared and the fetch authenticates with a service
account key in Actions secrets — a bigger change, not currently done.

**The decks are a deliberate exception, not a leak.** Their ids sit in committed
history and in already-deployed `schedule.json` versions, and both decks are
still world-readable by id (verified 2026-09-04: anonymous `export/pptx`
returns a real .pptx, ~150 KB each). That cannot be closed and **should not
be** — they are embedded in Veracross, the school's LMS, so link-sharing is
what makes them work. Do not propose restricting them.

So the model is: **decks are student-facing by design; the sheet is
teacher-side and stays unlisted.** The guards above protect the sheet, which is
the document holding the whole year's planning, drafts included. A stranger who
mines the repo history for deck ids gets 8th-grade Literature lesson plans that
the students already have through the LMS.

## Idea parked: writing board edits back to the sheet

Raised 2026-09-04, deferred. Today a board edit lives in `localStorage` only —
that device, that date — so an on-the-fly change in class never reaches the
sheet, the other devices, or the decks. Making it flow back is feasible; the
notes below are the parts that are already settled or already known to bite.

**Write to `Lesson Planner`, never `Master Data`.** Master Data's content
columns are formulas pointing back at Lesson Planner
(`=IF('Lesson Planner'!C39="","",…)`). Writing a value into one replaces the
formula and silently breaks that row for good. The target cell arithmetic is
already worked out in `tools/generate-lesson-planner-repair.ps1`: 18-row stride
per cycle block, fixed row offset per field, column = lesson number. With
`rotation.json` that makes *(period, date)* → cell a solved mapping.

**No write credential can live in the board.** It is a static page in a public
repo, so an endpoint URL or token in its JS is public — and this is *write*
access, unlike the sheet id. The workable shape is an Apps Script Web App on
the spreadsheet, with the board reading its endpoint and token from
`localStorage`, pasted in once per device. Nothing sensitive enters the repo.

**The decision that unblocks it: HP's three periods share one sheet row.** A, D
and G are one row per lesson, so an edit made during A period rewrites what D
and G will later read; the sheet cannot express "just for A". So: does a board
edit change all three, or do per-period tweaks stay board-only, or do only the
daily fields (`first5`, `materials`, `due`) write back? A conflict rule is
needed too — a board edit currently shadows the sheet on that device forever.

## Conventions

- **PowerShell 5.1 compatible.** No `&&` / `||`, no ternary, no `??`. The script
  runs under Windows PowerShell locally and `pwsh` on Linux in CI, so guard
  platform-specific calls (see the `ServicePointManager` try/catch).
- **Byte-identical JSON across PowerShell versions.** 5.1 and pwsh 7 disagree
  about indentation, non-ASCII escaping, and whether `& < > '` get escaped, so
  the same data serialised under each produces a different file and every run
  looks like a change. `-Compress` settles whitespace; the two regex passes
  after it settle the rest. Do not "simplify" them. See commit 070ee40.
- **`generatedAt` must not cause churn.** When only the clock moved, the old
  bytes are kept verbatim, so the field means "when the lesson data last
  changed". Verify after any change to the emit path: two consecutive runs must
  leave `schedule.json` byte-identical.
- **Fail safe, not silent.** The extractor checks the download is really a zip
  and aborts with the previous `schedule.json` intact, so a de-shared sheet
  cannot blank the board. Header drift throws rather than guessing.
- The run prints a per-record change report; that is how a sheet edit gets
  verified, since `-Compress` makes the diff itself unreadable.
- **A non-empty `warnings[]` fails the workflow run.** The `refresh` job emits
  each warning as an annotation and exports the count; a separate
  `report-warnings` job then fails, so a mistake in the sheet arrives as a
  failure email instead of sitting unread in a log. It is a *sibling* of
  `deploy`, not a gate on it — both need only `refresh` — because a warning
  means one lesson cannot display while the rest of the board is still fine and
  should still publish.
- **Commit messages** are imperative and sentence-case, describing the
  user-visible effect ("Shrink panel text to fit instead of scrolling"). The
  automated refresh uses `chore: refresh schedule from planning sheet`.
- `docs/robots.txt` asks search engines not to index the board.
- `design/Classroom Board.dc.html` is a design canvas artboard, not part of the
  deployed site. `docs/_ds/` is its design-system bundle.

## Testing without touching the live board

**Locally — covers everything the board does.** Neither step contacts GitHub or
deploys; the extractor only reads the sheet and writes `docs/data/schedule.json`.

```
./extract-schedule.ps1                              # rebuild the data
powershell -File tools/serve-board.ps1              # then open localhost:8099
```

A server is required: `index.html` loads its data with `fetch()`, which browsers
refuse over `file://`. Double-clicking `docs/index.html` gives a board with every
panel blank — looks like a data fault, isn't one.

Useful extractor flags for testing: `-Today <yyyy-MM-dd>` and
`-PublishLeadDays N` to exercise the horizon, and `-SiteDir <tmp>` to write the
JSON somewhere disposable instead of over `docs/data/`.

**In CI — covers what local runs cannot.** Three things only the runner
exercises: the byte-identical JSON logic under `pwsh` on Linux, the
`report-warnings` job, and the deploy itself. Push a branch and start the
workflow manually (`workflow_dispatch`); the Pages steps and the `deploy` job
are gated on `github.ref == 'refs/heads/main'`, so a branch run tests the
pipeline end to end and publishes nothing. It does commit `schedule.json` to
that branch, which is the point.

## Status: LIVE since 2026-09-04

The board runs on the planning sheet. [PR #1][pr1] merged as `961d89e`, and the
`main` push triggered the switchover run.

[pr1]: https://github.com/georgewestlund-hash/daily-holocron/pull/1

Verified on the live site, not just locally:

| Check | Result |
|---|---|
| CI run on `main` | success — `refresh` 14s, `deploy` 9s |
| `report-warnings` job | *skipped*, as designed: 0 warnings means its `if` is false |
| Byte-identical JSON under `pwsh` on Linux | "no change at all — file left byte-identical" |
| Sheet id in the public Actions log | masked to `sheet ...s52s` |
| Live `/data/schedule.json` | sheet-derived; no id, no export URL |
| Live `/data/deck-<id>.pptx` (both decks) | **404** — previously HTTP 200, ~160 KB each |

The `SHEET_ID` Actions variable exists. Repo variables are visible only to
accounts with repo access, so this is not the same exposure as putting the id
in the source.

### Merging cost two conflict resolutions, and would cost more if it had waited

`main`'s six-hourly job kept committing a deck-derived `schedule.json` while the
branch was replacing that pipeline, so the PR conflicted on a wholly derived
file — twice, once mid-session. Resolved each time by re-running the extractor
rather than hand-merging minified JSON: the sheet is the authority.

Nothing to guard against now — the old job *is* the new pipeline — but the shape
is worth remembering for any future long-lived branch that touches
`docs/data/schedule.json`.

### First 5 that "did not sync"

Reported once on the day of the switch; a hard refresh fixed it, and all 15
sheet rows carrying `First 5` were verified present in `schedule.json`. Two
distinct causes look identical from the classroom, and only one is a cache:

1. **Stale page.** The board fetches with `cache: 'no-store'`, but Pages' CDN
   caches `schedule.json` briefly and the page only loads it once. A hard
   refresh is the fix.
2. **`localStorage` wins.** `applyDeck` returns early when a stored value
   exists for that `(date, key)` — `if (saved !== null && saved !== '') return`.
   `first5`, `materials` and `due` were hand-typed for months before they became
   sheet-backed, so on any board where they were typed, the old text
   *permanently* shadows the sheet for that date. **A hard refresh will not fix
   this one.** The "Reload lesson" button does: it calls `applyDeck(true)`,
   which clears the stored value first.

### The student-facing channel is the decks, not the board

Worth being precise about, because it relocates a concern that was analysed in
the wrong place.

The [publication horizon](#publication-horizon--built-currently-off) exists so
unfinished planning is not published to students, and it is **off** because the
board's URL is not shared with anyone. That reasoning holds for the board.

But the decks **are** shared with students — they are embedded in Veracross —
and `sync-sheet-to-slides.gs` has no horizon at all. It pushes whatever Master
Data holds for whatever cycles are in scope. So the risk the horizon was built
for lives here instead:

- A half-drafted cycle synced early is **immediately visible to students** in
  the LMS.
- Worse, it is **effectively permanent**, because `ONLY_FILL_EMPTY_CELLS` is
  `true`: once a deck cell is filled the sheet can no longer correct it. Fixing
  the sheet updates the board and leaves the deck showing the draft.

The lever is scoping, not a horizon: set `ONLY_SEMESTER` and `ONLY_CYCLES` to
the cycle actually being taught, and run `preview` first. Syncing the whole
sheet in one go is what to avoid while later cycles are still drafts.

### Done (kept for context)

- Cycle 1–2 flattened content: repaired 2026-09-02 via
  `tools/apply-sheet-fixes.gs`, 71 of 71 cells verified. `H24` deliberately
  excluded — the deck parse there was worse than the curated text.
- The two colliding `Lesson` headers: renamed to `Lesson #` / `Lesson Plan`.
- Semester 1 cycle 3's seventh lesson: removed on 2026-09-02, restored the same
  day as a "correction", and on **2026-09-04 the removal turned out to have
  been right all along**. The four cycle-3 slots absent from the deck-derived
  skeleton are the four MAP testing periods, not lost data. `build-rotation.ps1`
  now models them explicitly and cycle 3 rebuilds as six lessons across seven
  meetings. See *A meeting is not a lesson*.

  Still to undo from the reversal: `Cycle Calendar!E4` is `7` and should be
  `6`, and `Lesson Planner!I38` is `Lesson 7` and should be greyed to `N/A`.

### The health check worth re-running

The sheet and the skeleton should be in exact 1:1 correspondence: every
planning cell maps to a real class meeting, and every meeting has a cell to
plan it in.

```
skeleton lesson keys : 218   (semester|cycle|lessonNumber|track)
Master Data rows     : 220   ← two too many, see below
phantom rows         : 2     sheet offers a slot the board can never show
orphan slots         : 0     board meeting with nowhere to plan it
```

218 is also an arithmetic check: lessons per cycle summed over both semesters
(54 + 55) × 2 tracks. It only balances if every cycle's lesson count is right,
cycle 3's **six** included.

**The two phantoms are S1 cycle 3 lesson 7, HP and EP** — left over from when
cycle 3 was thought to have seven lessons. They are harmless today because they
are blank, so nothing warns; the clean fix is `Cycle Calendar!E4` = 6, which
should drop the generated rows to 218. Until then this check reads 218 vs 220
by design, not by accident.

Re-run it after any school-calendar change — it is the fastest way to catch a
`ROT_DAYS`, calendar, `Cycle Calendar` or `$NON_INSTRUCTIONAL` disagreement.
Note that it was this check reading 220 vs 220 that made cycle 3's seventh
lesson look correct in the first place: the sheet had been edited to match the
skeleton, so the two sides were no longer independent. A balanced count is
evidence only when neither side was tuned to the other.

### Just how the data stands

As of 2026-09-04, Master Data is filled for **S1 cycles 1, 2 and 3** — 87 board
records of 440. `Due` is empty throughout, and `Bring` only for cycle 2 lessons
5–6. Blanks are handled correctly; nothing to fix there.

**Leftover placeholders: `First 5` says `TESTING` for S1 cycle 6 HP, lessons
1–5.** Five sheet rows, 15 board records. Worth knowing how this survived a
careful review: `TESTING` was pasted across cycles 3–6 while trying the slide
sync, then cleared — but the slide sync only reads `Lit. Focus`, `Lesson Plan`,
`Study` and `Homework`, so its preview never mentioned column E and the leftover
went unnoticed. **The sync's report is not a check on the whole sheet.** The
board is what surfaces all seven fields.

Cycle 3 is now planned in the sheet for both tracks, but the two are not
equivalent: HP's deck already held hand-written cycle-3 text that predates the
sheet and differs from it in substance, while EP's deck cycle 3 was empty and
gets filled by the sync. With `ONLY_FILL_EMPTY_CELLS` on, HP keeps its older
wording, so the two decks will show different plans for the same cycle until
that is resolved — either by clearing HP's cycle-3 cells, or by a scoped
overwrite run (`ONLY_SEMESTER: 1`, `ONLY_CYCLES: [3]`,
`ONLY_FILL_EMPTY_CELLS: false`).
