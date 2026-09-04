/**
 * Reformat the Classwork/Homework decks so every slide matches slides 2 & 3.
 *
 * TO RUN:  script.google.com > New project, paste this whole file, then
 *          1. Services (+) > Google Slides API > Add        [see note below]
 *          2. pick "preview" from the function dropdown and press Run
 *          3. read the execution log carefully
 *          4. pick "formatAllDecks" and press Run
 *
 * The Slides API service is optional. Without it everything still works
 * except the table resizing, which needs the advanced service to set column
 * widths. The script detects this itself and says so in the log.
 *
 * What it does to every slide from FIRST_TARGET_SLIDE onwards:
 *
 *   1. Deletes the class rows that slides 2 & 3 do not have. The rows to keep
 *      are read off slide 2, so this file needs no per-deck editing: the
 *      Honors deck keeps A, D and G, the ECP deck keeps E.
 *   2. Renames "Classwork" / "Class" to "CW" and "Homework" to "HW".
 *   3. Copies the fonts, colours, cell shading and alignment from the matching
 *      table on slide 2 onto every table.
 *   4. Resizes and repositions each table to the geometry of its slide 2
 *      counterpart, spreading the data columns evenly across the width.
 *   5. Rebuilds the title box as a copy of slide 2's, carrying over the cycle
 *      number, the date range, and S1 or S2 worked out from the slide's dates.
 *
 * NOTHING IS TYPED IN BY HAND. Every font, size, colour and measurement is
 * read off slide 2 at run time, so if you restyle slide 2 and run this again,
 * the rest of the deck follows.
 *
 * Safety: "preview" changes nothing at all - it only reports. "formatAllDecks"
 * saves a dated backup copy of each deck to your Drive before touching it, and
 * Slides' own File > Version history is still there as a second net.
 */

var CONFIG = {

  // The two decks. Add more ids here and they get the same treatment.
  PRESENTATION_IDS: [
    '1qk_g2tVKqYpposzORBzEtqEe3fsdVhcUfih14d4z46A',  // 8 Honors Literature
    '1VDpCkYyOhH1SllDQPwJxX81AXgIs8XzHlapY4RAvHaU'   // 8 ECP Literature
  ],

  STYLE_SOURCE_SLIDE: 2,    // 1-based. The slide everything is copied from.
  FIRST_TARGET_SLIDE: 4,    // 1-based. Slides 1-3 are left alone.

  // Which class rows each deck keeps. Normally this is read off slide 2, but
  // pinning it here is authoritative and removes any doubt. Set a deck to null
  // (or delete its line) to go back to reading it off slide 2. The log always
  // prints both, so a disagreement is visible.
  KEEP_PERIODS: {
    '1qk_g2tVKqYpposzORBzEtqEe3fsdVhcUfih14d4z46A': ['A', 'D', 'G'],   // Honors
    '1VDpCkYyOhH1SllDQPwJxX81AXgIs8XzHlapY4RAvHaU': ['E']              // ECP
  },

  MAKE_BACKUP_COPY:     true,   // copy the deck to Drive before editing
  DELETE_CLASS_ROWS:    true,   // drop period rows slide 2 does not have
  RENAME_CW_HW:         true,   // Classwork -> CW, Homework -> HW
  MATCH_TABLE_STYLE:    true,   // font family, size, weight, alignment
  MATCH_TABLE_GEOMETRY: true,   // position, width, column widths, row heights

  // Colour is deliberately OFF. Text colour and cell fill have to travel
  // together or not at all: slide 2 has white text on dark cells, and a cell
  // whose fill does not read back as a solid colour would otherwise end up
  // white-on-white. Turn both on together if you want the shading copied too.
  MATCH_TEXT_COLOUR: false,
  MATCH_CELL_FILL:   false,

  // Centre each table on the slide rather than copying slide 2's left edge.
  CENTER_TABLES: true,
  REBUILD_TITLE:        true,   // retitle from slide 2's box
  MATCH_COMMENTS_STYLE: true,   // font of the Comments/Important Notes box

  // A row whose first cell is blank is left alone by default: it is usually a
  // spare row rather than a class. Set true to delete those too.
  DELETE_UNLABELLED_ROWS: false,

  // Where the S1 / S2 in a rebuilt title comes from.
  //   'dates' - worked out from the slide's own date range (recommended: a
  //             handful of slides carry the wrong semester in their wording)
  //   'label' - whatever the slide already says, right or wrong
  // Either way a disagreement between the two is reported.
  SEMESTER_FROM: 'dates',

  // First month of semester 1. August..December are S1, January..May are S2.
  SEMESTER_1_FIRST_MONTH: 6,

  // Google rejects very narrow columns. Points.
  MIN_COLUMN_WIDTH: 32
};


/* ---------------------------------------------------------------- entry points */

/** Reports what would change. Modifies nothing. Run this first. */
function preview() {
  run_(true);
}

/** Does it for real, after taking a backup copy of each deck. */
function formatAllDecks() {
  run_(false);
}

/**
 * Diagnostic: prints the shape of every slide - its title, its tables and the
 * row labels it found. Use this if a slide comes out wrong and you want to see
 * what the script is actually looking at.
 */
function listSlides() {
  CONFIG.PRESENTATION_IDS.forEach(function (id) {
    var pres = SlidesApp.openById(id);
    say_('=== ' + pres.getName() + ' ===');
    pres.getSlides().forEach(function (slide, i) {
      var title = findTitleShape_(slide);
      say_('slide ' + (i + 1) + '  title: ' + (title ? oneLine_(title.getText().asString()) : '(none)'));
      slide.getTables().forEach(function (t) {
        say_('    ' + classifyTable_(t) + ' table, ' + t.getNumRows() + 'r x ' +
             t.getNumColumns() + 'c, col0 = [' + rowLabels_(t).join(' | ') + ']');
      });
    });
  });
}


/* ---------------------------------------------------------------- driver */

function run_(dryRun) {
  say_(dryRun ? '*** PREVIEW - nothing will be changed ***'
              : '*** LIVE RUN - decks will be modified ***');
  say_(hasAdvancedSlides_()
    ? 'Slides API service: enabled (tables will be resized)'
    : 'Slides API service: NOT enabled - skipping table resizing. ' +
      'Add it under Services (+) > Google Slides API to switch that on.');

  CONFIG.PRESENTATION_IDS.forEach(function (id) {
    try {
      processDeck_(id, dryRun);
    } catch (e) {
      say_('!! ' + id + ' FAILED: ' + e.message + '\n' + (e.stack || ''));
    }
  });
  say_('done.');
}

function processDeck_(id, dryRun) {
  var pres = SlidesApp.openById(id);
  say_('\n=== ' + pres.getName() + ' ===');

  var slides = pres.getSlides();
  var refIdx = CONFIG.STYLE_SOURCE_SLIDE - 1;
  if (slides.length <= refIdx) throw new Error('deck has no slide ' + CONFIG.STYLE_SOURCE_SLIDE);

  var ref = readReference_(slides[refIdx], id);
  say_('style source: slide ' + CONFIG.STYLE_SOURCE_SLIDE +
       '  title pattern: "' + oneLine_(ref.titleText) + '"');
  ['rotation', 'content'].forEach(function (kind) {
    var t = ref.tables[kind];
    if (!t) { say_('   no ' + kind + ' table found on slide ' + CONFIG.STYLE_SOURCE_SLIDE); return; }
    say_('   ' + kind + ' table: ' + Math.round(t.width) + 'pt wide, label column ' +
         Math.round(t.labelColWidth) + 'pt, at ' + Math.round(t.left) + ',' +
         Math.round(t.top) + ' (measured by ' + t.measuredBy + ')');
  });

  var pinned = CONFIG.KEEP_PERIODS ? CONFIG.KEEP_PERIODS[id] : null;
  say_('keep periods: slide ' + CONFIG.STYLE_SOURCE_SLIDE + ' says [' +
       ref.keepPeriods.join(', ') + ']' +
       (pinned ? ', CONFIG pins [' + pinned.join(', ') + '] - using CONFIG' : ''));
  if (pinned && pinned.length) {
    if (ref.keepPeriods.join(',') !== pinned.join(',')) {
      say_('   NOTE: those disagree. CONFIG wins, but slide ' + CONFIG.STYLE_SOURCE_SLIDE +
           ' is also where the fonts and table sizes come from, so run listSlides() ' +
           'to check the right table is being read.');
    }
    ref.keepPeriods = pinned;
  }

  if (!ref.keepPeriods.length) {
    throw new Error('no class rows found on slide ' + CONFIG.STYLE_SOURCE_SLIDE +
                    ' - cannot tell which rows to keep. Run listSlides() to see why, ' +
                    'or pin the list in CONFIG.KEEP_PERIODS.');
  }

  if (!dryRun && CONFIG.MAKE_BACKUP_COPY) {
    var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
    var copy = DriveApp.getFileById(id).makeCopy('BACKUP ' + stamp + ' - ' + pres.getName());
    say_('backup saved to your Drive: ' + copy.getName());
  }

  var geomReqs = [];
  for (var i = CONFIG.FIRST_TARGET_SLIDE - 1; i < slides.length; i++) {
    processSlide_(slides[i], i + 1, ref, dryRun, geomReqs, id);
  }

  if (!dryRun) {
    pres.saveAndClose();                      // flush the SlidesApp edits first
    flushGeometry_(id, geomReqs);
  }
}

/**
 * Sends the resize requests one table at a time. Batching the whole deck into
 * a single call would mean one table Google dislikes takes the other forty
 * with it.
 */
function flushGeometry_(id, geomReqs) {
  if (!geomReqs.length || !hasAdvancedSlides_()) return;
  var done = 0, failed = [];
  geomReqs.forEach(function (job) {
    try {
      Slides.Presentations.batchUpdate({ requests: job.reqs }, id);
      done++;
    } catch (e) {
      failed.push(job.label + ': ' + e.message);
    }
  });
  say_('resized ' + done + ' of ' + geomReqs.length + ' tables');
  if (failed.length) say_('   resize failed on: ' + failed.join(' | ') +
                          '  (everything else still applied)');
}


/* ---------------------------------------------------------------- one slide */

function processSlide_(slide, num, ref, dryRun, geomReqs, presId) {
  var tables = slide.getTables();
  var notes = [];

  if (!tables.length && !findTitleShape_(slide)) {
    say_('slide ' + num + ': empty, skipped');
    return;
  }

  tables.forEach(function (table) {
    var kind = classifyTable_(table);
    if (kind === 'unknown') { notes.push('left an unrecognised table alone'); return; }
    var refTable = ref.tables[kind];
    if (!refTable) { notes.push('no ' + kind + ' table on slide 2 to copy from'); return; }

    if (kind === 'rotation' && CONFIG.DELETE_CLASS_ROWS) {
      notes = notes.concat(deleteClassRows_(table, ref.keepPeriods, dryRun));
    }
    if (kind === 'content' && CONFIG.RENAME_CW_HW) {
      notes = notes.concat(renameHeaders_(table, dryRun));
    }
    if (CONFIG.MATCH_TABLE_STYLE && !dryRun) {
      styleTable_(table, refTable);
    }
    if (CONFIG.MATCH_TABLE_GEOMETRY) {
      notes = notes.concat(fitTable_(table, refTable, dryRun, geomReqs,
                                     'slide ' + num + ' ' + kind, presId));
    }
  });

  if (CONFIG.REBUILD_TITLE) {
    notes = notes.concat(rebuildTitle_(slide, ref, dryRun));
  }
  if (CONFIG.MATCH_COMMENTS_STYLE && !dryRun) {
    styleComments_(slide, ref);
  }

  say_('slide ' + num + ': ' + (notes.length ? notes.join('; ') : 'nothing to change'));
}


/* ---------------------------------------------------------------- reference */

/** Reads everything that gets copied out of the style-source slide. */
function readReference_(slide, presId) {
  var ref = { tables: {}, keepPeriods: [], titleText: '', title: null, comments: null };

  slide.getTables().forEach(function (table) {
    var kind = classifyTable_(table);
    if (kind === 'unknown' || ref.tables[kind]) return;
    ref.tables[kind] = readTableRef_(table, presId);
    if (kind === 'rotation') {
      rowLabels_(table).forEach(function (label) {
        var p = periodOf_(label);
        if (p && ref.keepPeriods.indexOf(p) < 0) ref.keepPeriods.push(p);
      });
    }
  });

  var title = findTitleShape_(slide);
  if (title) {
    ref.titleText = title.getText().asString();
    ref.title = {
      shape: title,
      style: captureTextStyle_(title.getText()),
      left: title.getLeft(), top: title.getTop(),
      width: title.getWidth(), height: title.getHeight()
    };
  }

  var comments = findCommentsShape_(slide);
  if (comments) ref.comments = { style: captureTextStyle_(comments.getText()) };

  return ref;
}

/** Captures the look of a table: three cell roles plus its measurements. */
function readTableRef_(table, presId) {
  var rows = table.getNumRows(), cols = table.getNumColumns();
  var r = {
    header: null, label: null, body: null,
    left: null, top: null, width: null,
    labelColWidth: null, headerRowHeight: null, bodyRowHeight: null,
    measuredBy: 'nothing'
  };

  for (var c = 0; c < cols && !r.header; c++) r.header = captureCell_(table, 0, c);
  for (var i = 1; i < rows; i++) {
    if (!r.label) r.label = captureCell_(table, i, 0);
    for (var j = 1; j < cols && !r.body; j++) r.body = captureCell_(table, i, j);
  }
  r.label = r.label || r.header;
  r.body = r.body || r.label;

  // A table's overall size is NOT in the page element the way a shape's is -
  // SlidesApp's getWidth() returns null for tables, and a null here used to
  // collapse every column to the minimum. The real widths live one level down,
  // per column, and only the Slides API exposes them.
  var geom = restTableGeometry_(presId, table.getObjectId());
  if (geom && geom.width) {
    r.left = geom.left;
    r.top = geom.top;
    r.width = geom.width;
    r.labelColWidth = geom.columnWidths[0];
    r.headerRowHeight = geom.rowHeights[0];
    r.bodyRowHeight = geom.rowHeights.length > 1 ? geom.rowHeights[1] : geom.rowHeights[0];
    r.measuredBy = 'Slides API';
  } else {
    r.left = safe_(function () { return table.getLeft(); });
    r.top = safe_(function () { return table.getTop(); });
    r.width = safe_(function () { return table.getWidth(); });
    r.labelColWidth = safe_(function () { return table.getColumn(0).getWidth(); });
    r.headerRowHeight = safe_(function () { return table.getRow(0).getMinimumHeight(); });
    r.bodyRowHeight = safe_(function () {
      return table.getRow(Math.min(1, rows - 1)).getMinimumHeight();
    });
    r.measuredBy = r.width ? 'SlidesApp' : 'nothing';
  }
  return r;
}


/* ------------------------------------------------- geometry via the Slides API */

var REST_CACHE = {};

function restPresentation_(id) {
  if (!hasAdvancedSlides_()) return null;
  if (!REST_CACHE.hasOwnProperty(id)) {
    try { REST_CACHE[id] = Slides.Presentations.get(id); }
    catch (e) { REST_CACHE[id] = null; }
  }
  return REST_CACHE[id];
}

/** Column widths, row heights and position of one table, all in points. */
function restTableGeometry_(presId, objectId) {
  var pres = restPresentation_(presId);
  if (!pres || !pres.slides) return null;

  for (var s = 0; s < pres.slides.length; s++) {
    var els = pres.slides[s].pageElements || [];
    for (var e = 0; e < els.length; e++) {
      if (els[e].objectId === objectId && els[e].table) return readRestTable_(els[e]);
    }
  }
  return null;
}

function readRestTable_(el) {
  var cols = [], rows = [], total = 0;

  (el.table.tableColumns || []).forEach(function (c) {
    var w = toPoints_(c.columnWidth);
    cols.push(w);
    total += w;
  });
  (el.table.tableRows || []).forEach(function (row) {
    rows.push(toPoints_(row.rowHeight));
  });

  var tr = el.transform || {};
  return {
    columnWidths: cols,
    rowHeights: rows,
    width: total,
    left: toPoints_({ magnitude: tr.translateX, unit: tr.unit }),
    top: toPoints_({ magnitude: tr.translateY, unit: tr.unit })
  };
}

/** Slide width in points, for centring. */
function restPageWidth_(presId) {
  var pres = restPresentation_(presId);
  if (!pres || !pres.pageSize) return 0;
  return toPoints_(pres.pageSize.width);
}

/** Slides measures in EMU almost everywhere. 1pt = 12700 EMU. */
function toPoints_(dim) {
  if (!dim || dim.magnitude === null || dim.magnitude === undefined) return 0;
  return dim.unit === 'EMU' ? dim.magnitude / 12700 : dim.magnitude;
}

/** Style of one cell, or null if the cell is merged away or empty. */
function captureCell_(table, row, col) {
  var cell = table.getCell(row, col);
  if (!cell || cell.getMergeState() === SlidesApp.CellMergeState.MERGED) return null;
  if (!cell.getText().asString().replace(/\s/g, '')) return null;
  return {
    text: captureTextStyle_(cell.getText()),
    fill: captureFill_(cell),
    contentAlignment: safe_(function () { return cell.getContentAlignment(); })
  };
}


/* ---------------------------------------------------------------- tables */

/**
 * rotation = the calendar grid, one row per class period.
 * content  = the CW / HW grid.
 */
function classifyTable_(table) {
  var labels = rowLabels_(table);
  var periods = 0, cwhw = 0;
  labels.forEach(function (l) {
    if (periodOf_(l)) periods++;
    if (HEADER_RENAMES.hasOwnProperty(l.toLowerCase())) cwhw++;
  });
  if (periods >= 2) return 'rotation';
  if (cwhw >= 1) return 'content';
  if (periods >= 1) return 'rotation';
  return 'unknown';
}

function rowLabels_(table) {
  var out = [];
  for (var i = 0; i < table.getNumRows(); i++) {
    var cell = table.getCell(i, 0);
    out.push(cell && cell.getMergeState() !== SlidesApp.CellMergeState.MERGED
      ? cell.getText().asString().trim() : '');
  }
  return out;
}

/**
 * "A", "A Period", "Extra C Period" -> the letter. Anything else -> null.
 * Deliberately strict: "DAY 1 (A-E)" is a day-type row, not a class row, and
 * must not be mistaken for one.
 */
function periodOf_(label) {
  var m = /^(?:extra\s+)?([A-G])(?:\s*period)?\s*$/i.exec(String(label).trim());
  return m ? m[1].toUpperCase() : null;
}

function deleteClassRows_(table, keep, dryRun) {
  var labels = rowLabels_(table), notes = [], dropped = [], failed = [];

  for (var i = labels.length - 1; i >= 0; i--) {
    var p = periodOf_(labels[i]);
    var isBlank = !labels[i];
    var drop = p ? keep.indexOf(p) < 0
                 : (isBlank && CONFIG.DELETE_UNLABELLED_ROWS && i > 0);
    if (!drop) continue;

    // In a live run the table has already shrunk; in a preview it has not.
    var remaining = dryRun ? table.getNumRows() - dropped.length : table.getNumRows();
    if (remaining <= 1) {
      failed.push('row ' + (i + 1) + ' (would empty the table)');
      continue;
    }

    dropped.push(p || 'blank row ' + (i + 1));
    if (!dryRun) {
      try {
        table.getRow(i).remove();
      } catch (e) {
        failed.push((p || 'row ' + (i + 1)) + ': ' + e.message);
        dropped.pop();
      }
    }
  }
  if (dropped.length) notes.push((dryRun ? 'would delete ' : 'deleted ') + dropped.reverse().join(', '));
  if (failed.length) notes.push('COULD NOT delete ' + failed.join(', ') + ' (merged cells?)');

  var blanks = labels.filter(function (l, i) { return !l && i > 0; }).length;
  if (blanks && !CONFIG.DELETE_UNLABELLED_ROWS) notes.push(blanks + ' unlabelled row(s) left alone');
  return notes;
}

var HEADER_RENAMES = {
  'classwork': 'CW', 'class work': 'CW', 'class': 'CW', 'cw': 'CW',
  'homework': 'HW', 'home work': 'HW', 'hw': 'HW'
};

function renameHeaders_(table, dryRun) {
  var changed = [];
  for (var r = 0; r < table.getNumRows(); r++) {
    for (var c = 0; c < table.getNumColumns(); c++) {
      var cell = table.getCell(r, c);
      if (!cell || cell.getMergeState() === SlidesApp.CellMergeState.MERGED) continue;
      var txt = cell.getText().asString().trim();
      var want = HEADER_RENAMES[txt.toLowerCase()];
      if (!want || want === txt) continue;   // exact match only: "No HW" is left alone
      changed.push(txt + ' -> ' + want);
      if (!dryRun) cell.getText().setText(want);
    }
  }
  return changed.length ? [(dryRun ? 'would rename ' : 'renamed ') + changed.join(', ')] : [];
}

function styleTable_(table, ref) {
  var rows = table.getNumRows(), cols = table.getNumColumns();
  for (var r = 0; r < rows; r++) {
    for (var c = 0; c < cols; c++) {
      var cell = table.getCell(r, c);
      if (!cell || cell.getMergeState() === SlidesApp.CellMergeState.MERGED) continue;
      applyCellStyle_(cell, r === 0 ? ref.header : (c === 0 ? ref.label : ref.body));
    }
  }
}

function applyCellStyle_(cell, role) {
  if (!role) return;
  applyTextStyle_(cell.getText(), role.text);
  if (role.fill && CONFIG.MATCH_CELL_FILL) {
    safe_(function () { setColor_(cell.getFill(), role.fill); });
  }
  if (role.contentAlignment) safe_(function () { cell.setContentAlignment(role.contentAlignment); });
}

/**
 * Puts the table where slide 2's is, at the same overall width, with the label
 * column kept at its slide 2 width and the remaining columns sharing what is
 * left evenly. That is the "fit in the space" part: a five-day slide gets wide
 * columns, a nine-day slide gets narrow ones, both filling the same box.
 */
function fitTable_(table, ref, dryRun, geomReqs, label, presId) {
  // Never resize against a width we could not actually measure: falling back
  // to the minimum would squash the table to a sliver.
  if (!ref.width || ref.width < CONFIG.MIN_COLUMN_WIDTH * 2) {
    return ['RESIZE SKIPPED: slide ' + CONFIG.STYLE_SOURCE_SLIDE + "'s " +
            'table width could not be read (got ' + ref.width + ')'];
  }

  if (!hasAdvancedSlides_()) {
    if (!dryRun) safe_(function () { table.setLeft(ref.left); table.setTop(ref.top); });
    return ['moved into place (enable the Slides API service to resize too)'];
  }

  var cols = table.getNumColumns(), rows = table.getNumRows();
  var labelW = Math.max(CONFIG.MIN_COLUMN_WIDTH, ref.labelColWidth || ref.width * 0.14);
  var dataW = cols > 1
    ? Math.max(CONFIG.MIN_COLUMN_WIDTH, (ref.width - labelW) / (cols - 1))
    : Math.max(CONFIG.MIN_COLUMN_WIDTH, ref.width);
  var total = cols > 1 ? labelW + dataW * (cols - 1) : dataW;
  var id = table.getObjectId();

  // Position is set as an absolute transform rather than with setLeft(): that
  // also forces the scale back to 1, so the table really is as wide as its
  // columns say it is. A table left at some other scale renders a different
  // width from the sum of its columns, which is what threw the centring off.
  var pageW = restPageWidth_(presId);
  var left = (CONFIG.CENTER_TABLES && pageW) ? (pageW - total) / 2 : ref.left;

  var reqs = [colReq_(id, [0], cols > 1 ? labelW : dataW)];
  if (cols > 1) reqs.push(colReq_(id, range_(1, cols), dataW));
  if (ref.headerRowHeight) reqs.push(rowReq_(id, [0], ref.headerRowHeight));
  if (ref.bodyRowHeight && rows > 1) reqs.push(rowReq_(id, range_(1, rows), ref.bodyRowHeight));
  reqs.push(transformReq_(id, left, ref.top));

  if (!dryRun) geomReqs.push({ label: label, reqs: reqs });
  return [(dryRun ? 'would fit ' : 'fitting ') + cols + ' columns into ' +
          Math.round(total) + 'pt (' + Math.round(dataW) + 'pt each), ' +
          (CONFIG.CENTER_TABLES && pageW ? 'centred' : 'placed') + ' at x=' +
          Math.round(left) + ' on a ' + Math.round(pageW) + 'pt slide'];
}

/** Absolute placement, scale normalised to 1. */
function transformReq_(id, left, top) {
  return {
    updatePageElementTransform: {
      objectId: id,
      applyMode: 'ABSOLUTE',
      transform: { scaleX: 1, scaleY: 1, translateX: left, translateY: top, unit: 'PT' }
    }
  };
}

function colReq_(id, indices, width) {
  return {
    updateTableColumnProperties: {
      objectId: id,
      columnIndices: indices,
      columnProperties: { columnWidth: { magnitude: width, unit: 'PT' } },
      fields: 'columnWidth'
    }
  };
}

function rowReq_(id, indices, height) {
  return {
    updateTableRowProperties: {
      objectId: id,
      rowIndices: indices,
      tableRowProperties: { minRowHeight: { magnitude: height, unit: 'PT' } },
      fields: 'minRowHeight'
    }
  };
}


/* ---------------------------------------------------------------- title box */

var CYCLE_RE = /cycle\s*#?\s*(\d+)/i;
var RANGE_RE = /(\d{1,2}\s*\/\s*\d{1,2})\s*(?:-|–|—|to)\s*(\d{1,2}\s*\/\s*\d{1,2})/;
var SEM_TOKEN_RE = /\bS([12])\b/;                          // "S1 Cycle 2"
var SEM_WORDS_RE = /\b(1st|2nd|first|second)\s+semester\b/i;  // "2nd Semester Template"

/**
 * Replaces the slide's title with a copy of slide 2's box, carrying over the
 * cycle number, dates and semester that were already on the slide.
 */
function rebuildTitle_(slide, ref, dryRun) {
  if (!ref.title) return ['no title box on slide 2 to copy'];

  var old = findTitleShape_(slide);
  if (!old) return ['no title found on this slide, left alone'];

  var oldText = old.getText().asString();
  var cycle = CYCLE_RE.exec(oldText);
  var range = RANGE_RE.exec(oldText);
  if (!cycle || !range) {
    return ['title has no cycle/date to carry over, left alone: "' + oneLine_(oldText) + '"'];
  }

  var notes = [];
  var sem = semesterOf_(oldText);
  if (sem.written && sem.dated && sem.written !== sem.dated) {
    notes.push('SEMESTER MISMATCH: slide says ' + sem.written + ', its dates say ' +
               sem.dated + ' - using ' + sem.label);
  }

  var next = buildTitle_(ref.titleText, cycle[1], range[0], sem.label);
  if (oneLine_(next) === oneLine_(oldText)) return notes;

  if (!dryRun) {
    var shape = slide.insertShape(ref.title.shape);
    shape.setLeft(ref.title.left).setTop(ref.title.top)
         .setWidth(ref.title.width).setHeight(ref.title.height);
    shape.getText().setText(next);
    applyTextStyle_(shape.getText(), ref.title.style);
    old.remove();
  }
  notes.push((dryRun ? 'would retitle' : 'retitled') + ': "' + oneLine_(oldText) +
             '" -> "' + oneLine_(next) + '"');
  return notes;
}

/**
 * Which semester a title belongs to, both as the slide words it and as its own
 * dates imply. CONFIG.SEMESTER_FROM picks which one wins; the caller reports
 * any disagreement, because a few slides carry the wrong wording - "1st
 * Semester Template: Cycle 8, 4/20 - 5/3" runs through April.
 */
function semesterOf_(title) {
  var written = null;
  var token = SEM_TOKEN_RE.exec(title);
  if (token) {
    written = 'S' + token[1];
  } else {
    var words = SEM_WORDS_RE.exec(title);
    if (words) written = /^(1st|first)$/i.test(words[1]) ? 'S1' : 'S2';
  }

  var dated = null;
  var range = RANGE_RE.exec(title);
  if (range) {
    var month = parseInt(String(range[1]).split('/')[0], 10);
    if (month >= 1 && month <= 12) {
      dated = month >= CONFIG.SEMESTER_1_FIRST_MONTH ? 'S1' : 'S2';
    }
  }

  return {
    label: CONFIG.SEMESTER_FROM === 'label' ? (written || dated) : (dated || written),
    written: written,
    dated: dated
  };
}

/**
 * Slide 2's title with its own cycle number, dates and semester swapped for
 * this slide's. Slide 2 already carries an "S1", so semester 2's slides get
 * that swapped to "S2" in place - which is what stops semester 2's cycle 1
 * from being indistinguishable from semester 1's.
 */
function buildTitle_(pattern, cycle, range, semester) {
  var out = pattern
    .replace(CYCLE_RE, function (whole, n) { return whole.replace(n, cycle); })
    .replace(RANGE_RE, range);

  if (!semester) return out;
  if (SEM_TOKEN_RE.test(out)) return out.replace(SEM_TOKEN_RE, semester);
  if (SEM_WORDS_RE.test(out)) return out.replace(SEM_WORDS_RE, semester);
  return out.replace(CYCLE_RE, function (whole) { return semester + ' ' + whole; });
}

/** The topmost text box that names a cycle, ignoring the Comments box. */
function findTitleShape_(slide) {
  var best = null;
  slide.getShapes().forEach(function (shape) {
    var txt = safe_(function () { return shape.getText().asString(); });
    if (!txt || !CYCLE_RE.test(txt)) return;
    if (/^\s*comments/i.test(txt)) return;
    if (txt.length > 200) return;               // a content blob, not a title
    if (!best || shape.getTop() < best.getTop()) best = shape;
  });
  return best;
}

function findCommentsShape_(slide) {
  var found = null;
  slide.getShapes().forEach(function (shape) {
    var txt = safe_(function () { return shape.getText().asString(); });
    if (txt && /^\s*comments\s*\/?\s*important/i.test(txt) && !found) found = shape;
  });
  return found;
}

function styleComments_(slide, ref) {
  if (!ref.comments) return;
  var box = findCommentsShape_(slide);
  if (box) applyTextStyle_(box.getText(), ref.comments.style);
}


/* ---------------------------------------------------------------- text style */

function captureTextStyle_(textRange) {
  if (!textRange) return null;
  var s = textRange.getTextStyle();
  return {
    fontFamily: safe_(function () { return s.getFontFamily(); }),
    fontSize:   safe_(function () { return s.getFontSize(); }),
    bold:       safe_(function () { return s.isBold(); }),
    italic:     safe_(function () { return s.isItalic(); }),
    underline:  safe_(function () { return s.isUnderline(); }),
    color:      safe_(function () { return captureColor_(s.getForegroundColor()); }),
    alignment:  safe_(function () {
      var ps = textRange.getParagraphs();
      return ps.length ? ps[0].getRange().getParagraphStyle().getParagraphAlignment() : null;
    })
  };
}

function applyTextStyle_(textRange, st) {
  if (!textRange || !st) return;
  var s = textRange.getTextStyle();
  if (st.fontFamily != null) safe_(function () { s.setFontFamily(st.fontFamily); });
  if (st.fontSize != null) safe_(function () { s.setFontSize(st.fontSize); });
  if (st.bold != null) safe_(function () { s.setBold(st.bold); });
  if (st.italic != null) safe_(function () { s.setItalic(st.italic); });
  if (st.underline != null) safe_(function () { s.setUnderline(st.underline); });
  if (st.color && CONFIG.MATCH_TEXT_COLOUR) safe_(function () { setColor_(s, st.color); });
  if (st.alignment) safe_(function () {
    textRange.getParagraphs().forEach(function (p) {
      p.getRange().getParagraphStyle().setParagraphAlignment(st.alignment);
    });
  });
}

function captureColor_(color) {
  if (!color) return null;
  try {
    if (color.getColorType() === SlidesApp.ColorType.RGB) {
      return { rgb: color.asRgbColor().asHexString() };
    }
    if (color.getColorType() === SlidesApp.ColorType.THEME) {
      return { theme: color.asThemeColor().getThemeColorType() };
    }
  } catch (e) {}
  return null;
}

/** target is a TextStyle (setForegroundColor) or a Fill (setSolidFill). */
function setColor_(target, c) {
  var setter = target.setForegroundColor ? 'setForegroundColor' : 'setSolidFill';
  target[setter](c.rgb || c.theme);
}

function captureFill_(cell) {
  try {
    var fill = cell.getFill();
    if (fill.getType() === SlidesApp.FillType.SOLID) {
      return captureColor_(fill.getSolidFill().getColor());
    }
  } catch (e) {}
  return null;   // no fill on the reference cell means "leave the fill alone"
}


/* ---------------------------------------------------------------- plumbing */

function hasAdvancedSlides_() {
  try { return typeof Slides !== 'undefined' && !!Slides.Presentations; }
  catch (e) { return false; }
}

/** Runs fn, swallowing the API errors that mixed or empty ranges throw. */
function safe_(fn) {
  try { return fn(); } catch (e) { return null; }
}

function range_(from, to) {
  var out = [];
  for (var i = from; i < to; i++) out.push(i);
  return out;
}

function oneLine_(s) {
  return String(s).replace(/\s+/g, ' ').trim();
}

function say_(msg) {
  console.log(msg);
}
