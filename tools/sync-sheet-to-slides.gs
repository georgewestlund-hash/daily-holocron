/**
 * Push Master Data into the Classwork/Homework decks.
 *
 * The planning sheet is the database. This copies three of its columns into
 * every deck slide's CW cell and one into the HW cell, keyed on
 * (semester, cycle, lesson #, track) - the same key the board's extractor uses.
 *
 *   Lit. Focus  ┐
 *   Lesson Plan ├─▶ CW cell for that lesson's column
 *   Study       ┘
 *   Homework     ──▶ HW cell for that lesson's column
 *
 * TO INSTALL:  open the planning sheet > Extensions > Apps Script, paste this
 *              whole file as a new script file, Save, then reload the sheet.
 *              A "Daily Holocron" menu appears next to Help.
 *
 * TO RUN:      Daily Holocron > Preview sync to slides     (changes nothing)
 *              read the log, then
 *              Daily Holocron > Sync to slides              (writes, after a backup)
 *
 * It must be bound to the sheet, not standalone: that is how it reads Master
 * Data without the sheet id appearing anywhere in this file. See "Keeping the
 * sheet unlisted" in CLAUDE.md. Set CONFIG.SPREADSHEET_ID only if you really
 * need a standalone script, and then keep that copy out of the repo.
 *
 * How a slide is matched to sheet rows
 *
 *   deck id      -> track      HP (Honors, A/D/G) or EP (ECP, E). CONFIG.TRACKS.
 *   slide title  -> semester   "S1"/"S2" if the title says so, else worked out
 *                              from the title's date range. A disagreement
 *                              between the two is reported.
 *   slide title  -> cycle      "Cycle 3" anywhere in the title.
 *   CW/HW table  -> lesson #   read off the table's own "Lesson N" header row,
 *                              never from column position. Cycles 3, 4, 5 and 8
 *                              are split across two slides, so the second slide
 *                              starts at Lesson 4 or 5 and column 1 is not
 *                              lesson 1.
 *
 * Safety
 *
 *   - "preview" writes nothing at all. It reports every cell it would change,
 *     with the old and new text, plus everything it could not match.
 *   - It only ever FILLS EMPTY CELLS. Anything already typed into a deck cell
 *     is kept, and the run says what the sheet would have put there instead.
 *     So this adds the lessons the decks are missing and changes nothing that
 *     is already written. CONFIG.ONLY_FILL_EMPTY_CELLS turns that off, and
 *     then the sheet wins wherever both sides have text.
 *   - A blank sheet cell NEVER blanks a slide cell either. Master Data is only
 *     filled for S1 cycles 1-2; the decks still hold hand-written text for
 *     cycle 3 and "No HW" notes elsewhere, and none of it is touched.
 *   - A dated backup copy of each deck is saved to Drive before any edit, and
 *     Slides' own File > Version history is a second net.
 *   - Cell character styling (font, size, colour) is captured before the write
 *     and reapplied, so a synced cell keeps the look it had.
 */

var CONFIG = {

  // Deck id -> which Master Data Track its rows come from.
  TRACKS: {
    '1qk_g2tVKqYpposzORBzEtqEe3fsdVhcUfih14d4z46A': 'HP',  // 8 Honors Literature
    '1VDpCkYyOhH1SllDQPwJxX81AXgIs8XzHlapY4RAvHaU': 'EP'   // 8 ECP Literature
  },

  MASTER_DATA_TAB: 'Master Data',

  // Leave empty. Only for a standalone script; a bound script reads the active
  // spreadsheet and needs no id. Do not commit a value here.
  SPREADSHEET_ID: '',

  WRITE_CW: true,
  WRITE_HW: true,

  // true (the default) fills only cells that are empty on the slide. Anything
  // already typed into the deck is kept, and the run reports what the sheet
  // would have put there instead.
  //
  // The consequence, and it is the whole trade-off: once a cell is filled the
  // sheet can no longer update it. Editing a lesson in Master Data will change
  // the board but not the deck. Set this false to make the sheet win every
  // time, which is what you want if the sheet is meant to be authoritative.
  ONLY_FILL_EMPTY_CELLS: true,

  // false (the default) means an empty sheet cell leaves the slide alone.
  // true means the sheet is absolutely authoritative and empty wipes the slide.
  // Ignored while ONLY_FILL_EMPTY_CELLS is true - nothing is ever cleared then.
  BLANK_SHEET_CLEARS_SLIDE: false,

  MAKE_BACKUP_COPY: true,      // copy each deck to Drive before editing
  PRESERVE_CELL_STYLE: true,   // reapply the cell's font/size/colour after the write

  // How the three sheet columns are laid out inside one CW cell. A field with
  // no content contributes nothing - no stray label, no blank line.
  CW_LABELS: {
    objective: 'Lit. Focus:',
    lesson:    'Lesson:',
    study:     'Study:'
  },
  CW_SECTION_GAP: '\n\n',   // between Lit. Focus, Lesson and Study
  HW_ITEM_GAP:    '\n\n',   // between homework lines, matching the decks today

  // Slides 1-3 of each deck are the cover and the two style-reference slides
  // that tools/format-slide-decks.gs copies from. 1 scans everything; nothing
  // without a CW/HW table is touched either way.
  FIRST_TARGET_SLIDE: 1,

  // First month of semester 1. August..December are S1, January..May are S2.
  // Same value as tools/format-slide-decks.gs.
  SEMESTER_1_FIRST_MONTH: 6,

  // Restrict the run, for testing. null = everything.
  // e.g. ONLY_SEMESTER: 1, ONLY_CYCLES: [2, 3]
  ONLY_SEMESTER: null,
  ONLY_CYCLES:   null
};


/* ---------------------------------------------------------------- entry points */

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('Daily Holocron')
      .addItem('Preview sync to slides', 'previewSync')
      .addItem('Sync to slides', 'syncToSlides')
      .addSeparator()
      .addItem('List deck slides (diagnostic)', 'listSlides')
      .addToUi();
  } catch (e) { /* not running against a UI */ }
}

/** Reports every cell it would change. Modifies nothing. Run this first. */
function previewSync() {
  run_(true);
}

/** Does it for real, after taking a backup copy of each deck. */
function syncToSlides() {
  run_(false);
}

/**
 * Diagnostic: prints what the script sees on every slide - the title it parsed,
 * the CW/HW table it found, and the column -> lesson map it built. Use this
 * when a slide is skipped and the reason is not obvious.
 */
function listSlides() {
  eachDeck_(function (id, track) {
    var pres = SlidesApp.openById(id);
    say_('=== ' + pres.getName() + '   track ' + track + ' ===');
    pres.getSlides().forEach(function (slide, i) {
      var t = readTitle_(slide);
      var table = findContentTable_(slide);
      var bits = ['slide ' + (i + 1)];
      bits.push(t.text ? ('S' + t.semester + ' cycle ' + t.cycle) : 'no title');
      if (!table) {
        bits.push('no CW/HW table');
      } else {
        var map = lessonColumns_(table);
        var rows = contentRows_(table);
        bits.push(table.getNumRows() + 'r x ' + table.getNumColumns() + 'c');
        bits.push('CW row ' + (rows.cw === null ? '-' : rows.cw + 1) +
                  ', HW row ' + (rows.hw === null ? '-' : rows.hw + 1));
        bits.push('lessons ' + describeMap_(map));
      }
      say_('  ' + bits.join('  |  '));
      if (t.text) { say_('      title: ' + oneLine_(t.text)); }
    });
    say_('');
  });
  flush_();
}


/* --------------------------------------------------------------------- driver */

function run_(dryRun) {
  LOG = [];
  say_(dryRun ? '*** PREVIEW - nothing will be changed ***'
              : '*** LIVE RUN - decks will be edited ***');

  var data = readMasterData_();
  say_('Master Data: ' + data.rowCount + ' rows, ' + data.filled +
       ' with content, keyed on semester|cycle|lesson|track');
  say_('');

  if (CONFIG.ONLY_FILL_EMPTY_CELLS) {
    say_('Filling empty slide cells only. Anything already in the deck is kept, ' +
         'and reported so you can see what the sheet would have put there.');
    say_('');
  }

  var totals = { written: 0, unchanged: 0, kept: 0, blank: 0, missing: 0, problems: 0 };

  eachDeck_(function (id, track) {
    processDeck_(id, track, data, dryRun, totals);
  });

  say_('');
  say_('TOTAL  ' + (dryRun ? 'would write ' : 'wrote ') + totals.written +
       ' cell(s), ' + totals.unchanged + ' already correct, ' +
       totals.kept + ' kept (slide already had text), ' +
       totals.blank + ' left alone (sheet blank), ' +
       totals.missing + ' with no sheet row, ' +
       totals.problems + ' problem(s)');
  if (totals.problems) {
    say_('Problems above are things the script could not do. Nothing else was affected.');
  }
  flush_();
}

function eachDeck_(fn) {
  Object.keys(CONFIG.TRACKS).forEach(function (id) { fn(id, CONFIG.TRACKS[id]); });
}

function processDeck_(id, track, data, dryRun, totals) {
  var pres = SlidesApp.openById(id);
  say_('=== ' + pres.getName() + '   track ' + track + ' ===');

  if (!dryRun && CONFIG.MAKE_BACKUP_COPY) {
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    DriveApp.getFileById(id).makeCopy('BACKUP ' + stamp + ' - ' + pres.getName());
    say_('  backup copy saved to Drive');
  }

  var slides = pres.getSlides();

  // Two passes over the deck, because "the sheet has a Lesson 7 but no column
  // holds it" can only be answered deck-wide. Cycles 3, 4, 5 and 8 are split
  // across two slides, so a lesson missing from one slide is normally just on
  // its sibling; the genuine gap is a lesson no slide in the deck claims.
  var coverage = {};
  for (var i = 0; i < slides.length; i++) {
    if (i + 1 < CONFIG.FIRST_TARGET_SLIDE) { continue; }
    var t = findContentTable_(slides[i]);
    if (!t) { continue; }
    var ti = readTitle_(slides[i]);
    if (!ti.cycle || !ti.semester) { continue; }
    var ck = ti.semester + '|' + ti.cycle;
    if (!coverage[ck]) { coverage[ck] = {}; }
    lessonColumns_(t).claimed.forEach(function (n) { coverage[ck][n] = true; });
  }

  for (var j = 0; j < slides.length; j++) {
    if (j + 1 < CONFIG.FIRST_TARGET_SLIDE) { continue; }
    processSlide_(slides[j], j + 1, track, data, dryRun, totals);
  }

  // Reported once per cycle, after the slides. This is how the known cycle-3
  // defect surfaces: the decks give cycle 3 six lesson columns where the
  // rotation gives it seven lessons. See CLAUDE.md, "The two inputs".
  Object.keys(coverage).forEach(function (ck) {
    var parts = ck.split('|');
    if (CONFIG.ONLY_SEMESTER !== null && Number(parts[0]) !== Number(CONFIG.ONLY_SEMESTER)) { return; }
    if (CONFIG.ONLY_CYCLES && CONFIG.ONLY_CYCLES.indexOf(Number(parts[1])) < 0) { return; }
    data.lessonsFor(parts[0], parts[1], track).forEach(function (lesson) {
      if (coverage[ck][lesson]) { return; }
      say_('  PROBLEM  S' + parts[0] + ' cycle ' + parts[1] + ' lesson ' + lesson +
           ': the sheet has content for it, but no slide in this deck has a ' +
           '"Lesson ' + lesson + '" column for that cycle. Add the column, or ' +
           'that lesson cannot be shown here.');
      totals.problems++;
    });
  });
  say_('');
}

function processSlide_(slide, num, track, data, dryRun, totals) {
  var table = findContentTable_(slide);
  if (!table) { return; }                      // cover slides, dividers, etc.

  var head = 'slide ' + num;
  var title = readTitle_(slide);
  if (!title.cycle) {
    say_('  ' + head + ': SKIPPED - no cycle number in the title (' +
         (title.text ? oneLine_(title.text) : 'no title box found') + ')');
    totals.problems++;
    return;
  }
  if (!title.semester) {
    say_('  ' + head + ': SKIPPED - cannot tell S1 from S2; the title says neither, ' +
         'and has no date range to work it out from (' + oneLine_(title.text) + ')');
    totals.problems++;
    return;
  }

  if (CONFIG.ONLY_SEMESTER !== null && Number(title.semester) !== Number(CONFIG.ONLY_SEMESTER)) { return; }
  if (CONFIG.ONLY_CYCLES && CONFIG.ONLY_CYCLES.indexOf(Number(title.cycle)) < 0) { return; }

  head += '  S' + title.semester + ' cycle ' + title.cycle;
  if (title.disagreement) {
    say_('  ' + head + ': NOTE - the title says ' + title.disagreement +
         '; going with ' + (title.semesterFrom === 'dates' ? 'the dates' : 'the label'));
  }

  var rows = contentRows_(table);
  if (rows.cw === null && rows.hw === null) {
    say_('  ' + head + ': SKIPPED - found a CW/HW table with neither a CW nor an HW row');
    totals.problems++;
    return;
  }

  var map = lessonColumns_(table);
  if (!map.count) {
    say_('  ' + head + ': SKIPPED - no "Lesson N" header cells in the CW/HW table, ' +
         'so there is no way to tell which column is which lesson');
    totals.problems++;
    return;
  }

  var notes = [];
  var lessons = [];
  Object.keys(map.byCol).forEach(function (c) { lessons.push(map.byCol[c]); });
  lessons.sort(function (a, b) { return a - b; });

  lessons.forEach(function (lesson) {
    var col = map.byLesson[lesson];
    var key = title.semester + '|' + title.cycle + '|' + lesson + '|' + track;
    var rec = data.byKey[key];

    if (!rec) {
      notes.push('L' + lesson + ': no Master Data row for ' + key);
      totals.missing++;
      return;
    }
    if (CONFIG.WRITE_CW && rows.cw !== null) {
      writeCell_(table, rows.cw, col, composeCW_(rec), 'L' + lesson + ' CW',
                 notes, dryRun, totals);
    }
    if (CONFIG.WRITE_HW && rows.hw !== null) {
      writeCell_(table, rows.hw, col, composeHW_(rec), 'L' + lesson + ' HW',
                 notes, dryRun, totals);
    }
  });

  say_('  ' + head + '  columns ' + describeMap_(map));
  notes.forEach(function (n) { say_('      ' + n); });
}


/* ----------------------------------------------------------------- cell write */

function writeCell_(table, row, col, want, label, notes, dryRun, totals) {
  var cell;
  try {
    cell = table.getCell(row, col);
  } catch (e) {
    notes.push(label + ': PROBLEM - no cell at row ' + (row + 1) + ' col ' + (col + 1));
    totals.problems++;
    return;
  }
  if (!cell || cell.getMergeState() === SlidesApp.CellMergeState.MERGED) {
    notes.push(label + ': PROBLEM - that cell is merged into another; unmerge it ' +
               'in the deck and run again');
    totals.problems++;
    return;
  }

  var have = cell.getText().asString();

  if (!want) {
    if (!CONFIG.BLANK_SHEET_CLEARS_SLIDE) {
      if (norm_(have)) { notes.push(label + ': sheet is blank, slide text left alone'); }
      totals.blank++;
      return;
    }
    if (!norm_(have)) { totals.unchanged++; return; }
  }

  if (norm_(have) === norm_(want)) { totals.unchanged++; return; }

  // Anything already typed into the deck outranks the sheet. Checked after the
  // equality test above so a cell that already matches reads as "correct"
  // rather than "protected" - otherwise a re-run would report every synced cell
  // as a cell it refused to touch.
  if (CONFIG.ONLY_FILL_EMPTY_CELLS && norm_(have)) {
    notes.push(label + ': slide already has text, left alone' +
               '\n           kept:  ' + oneLine_(have) +
               '\n           sheet: ' + oneLine_(want));
    totals.kept++;
    return;
  }

  notes.push(label + (dryRun ? ': would set' : ': set') +
             '\n           was:  ' + oneLine_(have) +
             '\n           now:  ' + oneLine_(want));
  totals.written++;
  if (dryRun) { return; }

  var text = cell.getText();
  var style = CONFIG.PRESERVE_CELL_STYLE ? captureTextStyle_(text) : null;
  text.setText(want);
  if (style) { applyTextStyle_(cell.getText(), style); }
}


/* ------------------------------------------------------------- text composition */

/**
 * Lit. Focus + Lesson Plan + Study, in that order, each behind its label, with
 * a blank line between sections. Only non-empty fields appear.
 */
function composeCW_(rec) {
  var parts = [];
  ['objective', 'lesson', 'study'].forEach(function (field) {
    var body = rec[field];
    if (!body) { return; }
    var label = CONFIG.CW_LABELS[field];
    parts.push(label ? (label + '\n' + body) : body);
  });
  return parts.join(CONFIG.CW_SECTION_GAP);
}

/** Homework, one sheet line per item, separated the way the decks do it. */
function composeHW_(rec) {
  if (!rec.homework) { return ''; }
  return rec.homework.split('\n').join(CONFIG.HW_ITEM_GAP);
}


/* -------------------------------------------------------------- Master Data */

/**
 * Reads Master Data into an index on semester|cycle|lesson|track.
 *
 * Columns are addressed by position, never by header name: C and H were both
 * originally headed "Lesson" - C the number, H the content - so a name-keyed
 * lookup silently collides and every row looks like it has content. The header
 * row is still validated against the same contract the board's extractor uses,
 * so a real layout change throws instead of being guessed at.
 */
function readMasterData_() {
  var ss = CONFIG.SPREADSHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    : SpreadsheetApp.getActive();
  if (!ss) {
    throw new Error('No spreadsheet. Bind this script to the planning sheet ' +
                    '(Extensions > Apps Script from the sheet itself).');
  }
  var sh = ss.getSheetByName(CONFIG.MASTER_DATA_TAB);
  if (!sh) {
    throw new Error('No "' + CONFIG.MASTER_DATA_TAB + '" tab in ' + ss.getName());
  }

  var values = sh.getDataRange().getValues();
  if (!values.length) { throw new Error(CONFIG.MASTER_DATA_TAB + ' is empty'); }

  // Column letter -> the names that column is allowed to have.
  var EXPECT = [
    ['A', 'Semester',   ['Semester']],
    ['B', 'Cycle',      ['Cycle']],
    ['C', 'Lesson #',   ['Lesson #', 'Lesson']],
    ['D', 'Track',      ['Track']],
    ['E', 'First 5',    ['First 5']],
    ['F', 'Bring',      ['Bring']],
    ['G', 'Lit. Focus', ['Lit. Focus']],
    ['H', 'Lesson Plan',['Lesson Plan', 'Lesson']],
    ['I', 'Study',      ['Study']],
    ['J', 'Homework',   ['Homework']],
    ['K', 'Due',        ['Due']]
  ];
  var header = values[0], bad = [];
  EXPECT.forEach(function (spec, i) {
    var actual = String(header[i] === undefined ? '' : header[i]).trim();
    if (spec[2].indexOf(actual) < 0) {
      bad.push(spec[0] + '1 is "' + actual + '", expected ' +
               spec[2].map(function (n) { return '"' + n + '"'; }).join(' or '));
    }
  });
  if (bad.length) {
    throw new Error('"' + CONFIG.MASTER_DATA_TAB + '" layout has changed - ' +
                    'refusing to guess. ' + bad.join('; '));
  }

  var byKey = {}, byCycle = {}, filled = 0;
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var semester = digits_(row[0]);
    var cycle    = digits_(row[1]);
    var lesson   = digits_(row[2]);
    var track    = String(row[3] === undefined ? '' : row[3]).trim().toUpperCase();
    if (!semester || !cycle || !lesson || !track) { continue; }

    var rec = {
      objective: cell_(row[6]),
      lesson:    cell_(row[7]),
      study:     cell_(row[8]),
      homework:  cell_(row[9]),
      row:       r + 1
    };
    var has = !!(rec.objective || rec.lesson || rec.study || rec.homework);
    if (has) { filled++; }

    var key = semester + '|' + cycle + '|' + lesson + '|' + track;
    byKey[key] = rec;

    if (has) {
      var ck = semester + '|' + cycle + '|' + track;
      if (!byCycle[ck]) { byCycle[ck] = []; }
      byCycle[ck].push(Number(lesson));
    }
  }

  return {
    byKey: byKey,
    rowCount: values.length - 1,
    filled: filled,
    /** The lesson numbers this cycle actually has content for. */
    lessonsFor: function (semester, cycle, track) {
      var list = byCycle[semester + '|' + cycle + '|' + track];
      if (!list) { return []; }
      return list.slice().sort(function (a, b) { return a - b; });
    }
  };
}

/** Trims, and normalises the line endings a sheet cell can carry. */
function cell_(v) {
  if (v === undefined || v === null) { return ''; }
  var s = String(v).replace(/\r\n?/g, '\n');
  s = s.split('\n').map(function (l) { return l.replace(/\s+$/, ''); })
       .join('\n').replace(/\n{3,}/g, '\n\n');
  return s.replace(/^\s+|\s+$/g, '');
}

function digits_(v) {
  var m = /(\d+)/.exec(String(v === undefined ? '' : v));
  return m ? m[1] : '';
}


/* ----------------------------------------------------------- slide inspection */

var CW_LABELS_IN = { 'cw': 1, 'classwork': 1, 'class work': 1, 'class': 1 };
var HW_LABELS_IN = { 'hw': 1, 'homework': 1, 'home work': 1 };

/**
 * The CW/HW table on a slide, or null. Identified by a row label, not by
 * position: on some slides the content table is the first shape, on others the
 * last. A rotation table (two or more single-letter period rows) is never it.
 */
function findContentTable_(slide) {
  var tables = slide.getTables(), best = null;
  for (var i = 0; i < tables.length; i++) {
    var rows = contentRows_(tables[i]);
    if (rows.cw === null && rows.hw === null) { continue; }
    if (rows.periods >= 2) { continue; }        // a rotation grid
    if (!best) { best = tables[i]; }
  }
  return best;
}

/** Row indexes of the CW and HW rows, plus how many period rows there are. */
function contentRows_(table) {
  var out = { cw: null, hw: null, periods: 0 };
  for (var r = 0; r < table.getNumRows(); r++) {
    var label = rowLabel_(table, r).toLowerCase();
    if (out.cw === null && CW_LABELS_IN[label]) { out.cw = r; }
    if (out.hw === null && HW_LABELS_IN[label]) { out.hw = r; }
    if (/^(?:extra\s+)?[a-g](?:\s*period)?$/i.test(label)) { out.periods++; }
  }
  return out;
}

function rowLabel_(table, r) {
  try {
    var cell = table.getCell(r, 0);
    if (!cell || cell.getMergeState() === SlidesApp.CellMergeState.MERGED) { return ''; }
    return cell.getText().asString().replace(/\s+/g, ' ').replace(/^ | $/g, '');
  } catch (e) { return ''; }
}

/**
 * Which column holds which lesson, read off the table's own "Lesson N" cells.
 *
 * Never inferred from position. Cycles 3, 4, 5 and 8 are split over two slides,
 * and the second slide's first data column is Lesson 4 or Lesson 5.
 *
 * Returns { byCol, byLesson, count, claimed, dupes }. `claimed` is every lesson
 * number this table mentions, including any it could not map to a column;
 * processDeck_ unions it across the whole deck so that "this lesson is on the
 * cycle's other slide" is not mistaken for "this lesson has nowhere to go".
 */
function lessonColumns_(table) {
  var byCol = {}, byLesson = {}, dupes = [], claimed = [];
  for (var r = 0; r < table.getNumRows(); r++) {
    var label = rowLabel_(table, r).toLowerCase();
    if (CW_LABELS_IN[label] || HW_LABELS_IN[label]) { continue; }   // a data row
    for (var c = 1; c < table.getNumColumns(); c++) {
      var txt;
      try {
        var cell = table.getCell(r, c);
        if (!cell || cell.getMergeState() === SlidesApp.CellMergeState.MERGED) { continue; }
        txt = cell.getText().asString();
      } catch (e) { continue; }
      var m = /^\s*lesson\s*#?\s*(\d+)\s*$/i.exec(txt.replace(/\s+/g, ' '));
      if (!m) { continue; }
      var n = Number(m[1]);
      if (claimed.indexOf(n) < 0) { claimed.push(n); }
      if (byCol.hasOwnProperty(c)) {
        if (byCol[c] !== n) { dupes.push('col ' + (c + 1) + ' says both Lesson ' + byCol[c] + ' and Lesson ' + n); }
        continue;
      }
      if (byLesson.hasOwnProperty(n)) {
        dupes.push('Lesson ' + n + ' appears in columns ' + (byLesson[n] + 1) + ' and ' + (c + 1));
        continue;
      }
      byCol[c] = n;
      byLesson[n] = c;
    }
  }
  claimed.sort(function (a, b) { return a - b; });
  return { byCol: byCol, byLesson: byLesson, count: Object.keys(byCol).length,
           claimed: claimed, dupes: dupes };
}

function describeMap_(map) {
  var cols = Object.keys(map.byCol).sort(function (a, b) { return a - b; });
  if (!cols.length) { return '(none found)'; }
  var bits = cols.map(function (c) { return (Number(c) + 1) + '=L' + map.byCol[c]; });
  if (map.dupes.length) { bits.push('AMBIGUOUS: ' + map.dupes.join('; ')); }
  return bits.join(' ');
}

/**
 * The slide's cycle title, and the semester and cycle read out of it.
 *
 * Titles are not uniform. The Honors deck writes
 *   "8 Honors Lit. - Westlund – S1 - Cycle 3:  9/8 - 9/18"
 * and the ECP deck writes "8 EP Lit. - Westlund - S1 Cycle 3: ..." - or, for
 * its first two slides, no semester at all. So the semester comes from the
 * date range when the label does not give one, and the two are cross-checked
 * when both exist.
 */
function readTitle_(slide) {
  var out = { text: '', cycle: '', semester: '', semesterFrom: '', disagreement: '' };
  var shape = findTitleShape_(slide);
  if (!shape) { return out; }
  out.text = shape.getText().asString();

  var flat = out.text.replace(/\s+/g, ' ');
  var mc = /cycle\s*#?\s*(\d+)/i.exec(flat);
  if (mc) { out.cycle = mc[1]; }

  var labelled = '';
  var ms = /(?:^|[^A-Za-z])S\s*([12])(?![0-9A-Za-z])/.exec(flat);
  if (ms) { labelled = ms[1]; }

  var fromDates = semesterFromRange_(flat);

  if (fromDates) { out.semester = fromDates; out.semesterFrom = 'dates'; }
  else if (labelled) { out.semester = labelled; out.semesterFrom = 'label'; }

  if (labelled && fromDates && labelled !== fromDates) {
    out.disagreement = 'S' + labelled + ' but its dates fall in S' + fromDates;
  }
  return out;
}

/** "9/8 - 9/18" -> the semester its first month falls in, or ''. */
function semesterFromRange_(flat) {
  var m = /(\d{1,2})\s*\/\s*(\d{1,2})\s*[-–—]\s*(\d{1,2})\s*\/\s*(\d{1,2})/.exec(flat);
  if (!m) { return ''; }
  var month = Number(m[1]);
  if (month < 1 || month > 12) { return ''; }
  return month >= CONFIG.SEMESTER_1_FIRST_MONTH ? '1' : '2';
}

/**
 * The text box that names a cycle. The Comments/Important Notes box is skipped
 * explicitly - on several slides it sits above the title.
 */
function findTitleShape_(slide) {
  var shapes = slide.getShapes(), best = null;
  for (var i = 0; i < shapes.length; i++) {
    var txt;
    try { txt = shapes[i].getText().asString(); } catch (e) { continue; }
    if (!txt) { continue; }
    if (/comments\s*\/?\s*important/i.test(txt)) { continue; }
    if (!/cycle\s*#?\s*\d+/i.test(txt)) { continue; }
    if (!best || shapes[i].getTop() < best.getTop()) { best = shapes[i]; }
  }
  return best;
}


/* ------------------------------------------------------------------- styling */

/** The character style at the start of a range, so a rewrite can keep it. */
function captureTextStyle_(textRange) {
  var s;
  try { s = textRange.getTextStyle(); } catch (e) { return null; }
  if (!s) { return null; }
  var out = {};
  try { out.family = s.getFontFamily(); } catch (e) {}
  try { out.size = s.getFontSize(); } catch (e) {}
  try { out.bold = s.isBold(); } catch (e) {}
  try { out.italic = s.isItalic(); } catch (e) {}
  try { out.underline = s.isUnderline(); } catch (e) {}
  try {
    var c = s.getForegroundColor();
    if (c && c.getColorType() === SlidesApp.ColorType.RGB) { out.rgb = c.asRgbColor().asHexString(); }
  } catch (e) {}
  return out;
}

function applyTextStyle_(textRange, st) {
  if (!st) { return; }
  var s;
  try { s = textRange.getTextStyle(); } catch (e) { return; }
  if (!s) { return; }
  try { if (st.family) { s.setFontFamily(st.family); } } catch (e) {}
  try { if (st.size) { s.setFontSize(st.size); } } catch (e) {}
  try { if (st.bold !== undefined && st.bold !== null) { s.setBold(st.bold); } } catch (e) {}
  try { if (st.italic !== undefined && st.italic !== null) { s.setItalic(st.italic); } } catch (e) {}
  try { if (st.underline !== undefined && st.underline !== null) { s.setUnderline(st.underline); } } catch (e) {}
  try { if (st.rgb) { s.setForegroundColor(st.rgb); } } catch (e) {}
}


/* ------------------------------------------------------------------ reporting */

var LOG = [];

function say_(line) {
  LOG.push(line);
  Logger.log(line);
}

/**
 * Apps Script's execution log truncates. Writing the report to a Drive text
 * file as well means a long run is still readable in full.
 */
function flush_() {
  var body = LOG.join('\n');
  try {
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    DriveApp.createFile('Holocron slide sync ' + stamp + '.txt', body, MimeType.PLAIN_TEXT);
    Logger.log('Full report also saved to Drive as "Holocron slide sync ' + stamp + '.txt"');
  } catch (e) {
    Logger.log('(could not save the report to Drive: ' + e.message + ')');
  }
  try {
    SpreadsheetApp.getActive().toast('Slide sync finished - see the execution log',
                                     'Daily Holocron', 8);
  } catch (e) {}
}

/** Collapses newlines so a multi-line cell fits one log line. */
function oneLine_(s) {
  var t = String(s === undefined || s === null ? '' : s)
    .replace(/\n/g, ' ⏎ ').replace(/\s+/g, ' ').replace(/^ | $/g, '');
  return t.length > 160 ? t.slice(0, 157) + '...' : (t || '(empty)');
}

/**
 * Comparison form: enough normalising that a stray trailing space is not
 * reported as a change, but blank lines are preserved. The gap between CW's
 * three sections, and between HW items, is doing visible work in these cells,
 * so collapsing it here would report a cell as already correct when only its
 * spacing is wrong.
 */
function norm_(s) {
  var VT = String.fromCharCode(11);   // Slides returns a soft line break as \v
  return String(s === undefined || s === null ? '' : s)
    .split(VT).join('\n')
    .replace(/\r\n?/g, '\n')
    .replace(/[\u200b\u00a0]/g, ' ')
    .replace(/[ \t]+/g, ' ')
    .split('\n').map(function (l) { return l.replace(/^ | $/g, ''); })
    .join('\n').replace(/\n{3,}/g, '\n\n').replace(/^\s+|\s+$/g, '');
}
