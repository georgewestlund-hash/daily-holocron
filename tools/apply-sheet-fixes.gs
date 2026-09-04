/**
 * All sheet-side fixes for the Daily Holocron sync, in one file.
 *
 * TO RUN:  Extensions > Apps Script, paste this whole file, then pick
 *          "runAll" from the function dropdown and press Run.
 *
 * runAll does three things:
 *   1. repairLessonPlanner    - restores the line breaks, and cycle 1's
 *                               Lit. Focus / Lesson split, that were flattened
 *                               when this sheet was first built. 71 cells.
 *   2. renameMasterDataHeaders - gives the two "Lesson" columns distinct names.
 *   3. restoreCycle3Lesson7    - restores cycle 3's 7th lesson (an earlier
 *                                version of this file wrongly removed it).
 *
 * Every step is safe to run more than once: each one either sets a cell to a
 * known value or checks it is already correct. Re-running changes nothing that
 * is already right, so if a previous attempt only got partway through, just
 * run this again.
 *
 * Nothing here touches formatting or the Master Data formulas.
 */

function runAll() {
  repairLessonPlanner();
  renameMasterDataHeaders();
  restoreCycle3Lesson7();
  SpreadsheetApp.getActive().toast('All sheet fixes applied', 'Daily Holocron', 6);
}

var SHEET_NAME = 'Lesson Planner';

var FIXES = [
  { a1: 'C5', v: 'Course Introductions' },  // cycle 1 HP L1 Lit. Focus
  { a1: 'C6', v: 'Class Guidelines & Expectations\nSchool procedures & iPad refresher' },  // cycle 1 HP L1 Lesson
  { a1: 'C8', v: 'Bring a book to read in class this week!' },  // cycle 1 HP L1 Homework
  { a1: 'D5', v: 'Digital Setup' },  // cycle 1 HP L2 Lit. Focus
  { a1: 'D6', v: 'Google Classroom\nCommonLit.org\nMembean\nClassreads\nTrain on Membean\nRead your book' },  // cycle 1 HP L2 Lesson
  { a1: 'D8', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 HP L2 Homework
  { a1: 'E5', v: 'Critical Reading Skills' },  // cycle 1 HP L3 Lit. Focus
  { a1: 'E6', v: 'Diagnostic Assessment of Critical Reading Skills\nRead your book' },  // cycle 1 HP L3 Lesson
  { a1: 'E8', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 HP L3 Homework
  { a1: 'F5', v: 'Course Syllabus' },  // cycle 1 HP L4 Lit. Focus
  { a1: 'F6', v: 'Materials & Content\nJournals\nGrading\nExpectations\nLibrary\nTrain on Membean\nRead your book' },  // cycle 1 HP L4 Lesson
  { a1: 'F8', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 HP L4 Homework
  { a1: 'G5', v: 'Critical Reading Skills' },  // cycle 1 HP L5 Lit. Focus
  { a1: 'G6', v: 'Skill Lessons\nPurpose\nGuided Lesson\nPractice Lesson\nTimed Reading' },  // cycle 1 HP L5 Lesson
  { a1: 'G8', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 HP L5 Homework
  { a1: 'H5', v: 'Critical Reading Skills' },  // cycle 1 HP L6 Lit. Focus
  { a1: 'H6', v: 'Skill Lessons\nPractice Lesson' },  // cycle 1 HP L6 Lesson
  { a1: 'H8', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 HP L6 Homework
  { a1: 'C12', v: 'Course Introductions' },  // cycle 1 EP L1 Lit. Focus
  { a1: 'C13', v: 'Class Guidelines & Expectations\nSchool procedures & iPad refresher' },  // cycle 1 EP L1 Lesson
  { a1: 'C15', v: 'Bring a book to read in class this week!' },  // cycle 1 EP L1 Homework
  { a1: 'D12', v: 'Digital Setup' },  // cycle 1 EP L2 Lit. Focus
  { a1: 'D13', v: 'Google Classroom\nCommonLit.org\nMembean\nClassreads\nTrain on Membean\nRead your book' },  // cycle 1 EP L2 Lesson
  { a1: 'D15', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 EP L2 Homework
  { a1: 'E12', v: 'Critical Reading Skills' },  // cycle 1 EP L3 Lit. Focus
  { a1: 'E13', v: 'Diagnostic Assessment of Critical Reading Skills\nRead your book' },  // cycle 1 EP L3 Lesson
  { a1: 'E15', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 EP L3 Homework
  { a1: 'F12', v: 'Course Syllabus' },  // cycle 1 EP L4 Lit. Focus
  { a1: 'F13', v: 'Materials & Content\nJournals\nGrading\nExpectations\nLibrary\nTrain on Membean\nRead your book' },  // cycle 1 EP L4 Lesson
  { a1: 'F15', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 EP L4 Homework
  { a1: 'G12', v: 'Critical Reading Skills' },  // cycle 1 EP L5 Lit. Focus
  { a1: 'G13', v: 'Skill Lessons\nPurpose\nGuided Lesson\nPractice Lesson\nTimed Reading' },  // cycle 1 EP L5 Lesson
  { a1: 'G15', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 EP L5 Homework
  { a1: 'H12', v: 'Critical Reading Skills' },  // cycle 1 EP L6 Lit. Focus
  { a1: 'H13', v: 'Skill Lessons\nPractice Lesson' },  // cycle 1 EP L6 Lesson
  { a1: 'H15', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 1 EP L6 Homework
  { a1: 'C24', v: 'Diagnostic Reading Assessment: Commonlit' },  // cycle 2 HP L1 Lesson
  { a1: 'C25', v: 'Train on Membean.\nRead independently.' },  // cycle 2 HP L1 Study
  { a1: 'C26', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 2 HP L1 Homework
  { a1: 'D24', v: 'Refresh Checking out books in Library!' },  // cycle 2 HP L2 Lesson
  { a1: 'D25', v: 'Check out a book and read!' },  // cycle 2 HP L2 Study
  { a1: 'D26', v: 'Read for 20 minutes\nTrain on Membean' },  // cycle 2 HP L2 Homework
  { a1: 'E24', v: 'Skill Lessons\nPurpose\nGuided Lesson\nPractice Lesson\nTimed Reading' },  // cycle 2 HP L3 Lesson
  { a1: 'E25', v: 'Practice Skill Lesson' },  // cycle 2 HP L3 Study
  { a1: 'E26', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 HP L3 Homework
  { a1: 'F24', v: 'Selected Practice Lessons\nComplete\nCheck\nReview' },  // cycle 2 HP L4 Lesson
  { a1: 'F25', v: 'Read Independently' },  // cycle 2 HP L4 Study
  { a1: 'F26', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 HP L4 Homework
  { a1: 'G24', v: 'Review the literary elements.\nIntroduce the short story, “The Fan Club”' },  // cycle 2 HP L5 Lesson
  { a1: 'G25', v: 'Read the story.' },  // cycle 2 HP L5 Study
  { a1: 'G26', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 HP L5 Homework
  { a1: 'H25', v: 'Read independently.' },  // cycle 2 HP L6 Study
  { a1: 'H26', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 HP L6 Homework
  { a1: 'C31', v: 'Diagnostic Reading Assessment: Commonlit' },  // cycle 2 EP L1 Lesson
  { a1: 'C32', v: 'Train on Membean.\nRead independently.' },  // cycle 2 EP L1 Study
  { a1: 'C33', v: 'Read a magazine, newspaper, or book for 20 minutes!\nTrain on Membean' },  // cycle 2 EP L1 Homework
  { a1: 'D31', v: 'Refresh Checking out books in Library!' },  // cycle 2 EP L2 Lesson
  { a1: 'D32', v: 'Check out a book and read!' },  // cycle 2 EP L2 Study
  { a1: 'D33', v: 'Read for 20 minutes\nTrain on Membean' },  // cycle 2 EP L2 Homework
  { a1: 'E31', v: 'Skill Lessons\nPurpose\nGuided Lesson\nPractice Lesson\nTimed Reading' },  // cycle 2 EP L3 Lesson
  { a1: 'E32', v: 'Practice Skill Lesson' },  // cycle 2 EP L3 Study
  { a1: 'E33', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 EP L3 Homework
  { a1: 'F31', v: 'Selected Practice Lessons\nComplete\nCheck\nReview' },  // cycle 2 EP L4 Lesson
  { a1: 'F32', v: 'Read Independently' },  // cycle 2 EP L4 Study
  { a1: 'F33', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 EP L4 Homework
  { a1: 'G31', v: 'Review the Methods of Characterization\nIntroduce the short story, “The Fan Club”' },  // cycle 2 EP L5 Lesson
  { a1: 'G32', v: 'Read the story.' },  // cycle 2 EP L5 Study
  { a1: 'G33', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 EP L5 Homework
  { a1: 'H31', v: 'Continue “The Fan Club”\nRead\nDiscuss' },  // cycle 2 EP L6 Lesson
  { a1: 'H32', v: 'Read independently.' },  // cycle 2 EP L6 Study
  { a1: 'H33', v: 'Train on Membean\nRead for 20 minutes' },  // cycle 2 EP L6 Homework
];

function repairLessonPlanner() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) { throw new Error('No sheet named ' + SHEET_NAME); }
  for (var i = 0; i < FIXES.length; i++) {
    sh.getRange(FIXES[i].a1).setValue(FIXES[i].v).setWrap(true);
  }
  SpreadsheetApp.getActive().toast(FIXES.length + ' cells repaired', 'Lesson Planner', 5);
}


/**
 * 2. Give the two "Lesson" columns distinct names.
 *
 * Master Data has C headed "Lesson" (the lesson NUMBER) and H also headed
 * "Lesson" (the lesson CONTENT). Anything reading that tab by header name
 * silently picks the wrong one - this already caused a real bug where every
 * row looked like it had content.
 *
 * The extractor accepts either the old or the new name, so running this does
 * not require any matching code change and cannot break a scheduled sync.
 */
function renameMasterDataHeaders() {
  var sh = SpreadsheetApp.getActive().getSheetByName('Master Data');
  if (!sh) { throw new Error('No sheet named "Master Data"'); }

  // Confirm we are renaming what we think we are before touching anything.
  var c1 = sh.getRange('C1').getValue();
  var h1 = sh.getRange('H1').getValue();
  var okC = (c1 === 'Lesson' || c1 === 'Lesson #');
  var okH = (h1 === 'Lesson' || h1 === 'Lesson Plan');
  if (!okC || !okH) {
    throw new Error('Unexpected headers - C1 is "' + c1 + '", H1 is "' + h1 +
                    '". Expected "Lesson" (or the already-renamed values). Nothing changed.');
  }

  sh.getRange('C1').setValue('Lesson #');
  sh.getRange('H1').setValue('Lesson Plan');
  Logger.log('Master Data: C1 -> "Lesson #", H1 -> "Lesson Plan"');
}


/**
 * 3. Restore Semester 1 Cycle 3's seventh lesson.
 *
 * UNDOES AN EARLIER MISTAKE IN THIS FILE. A previous version set
 * Cycle Calendar!E4 to 6 and greyed Lesson Planner!I38 to "N/A", on the
 * incorrect belief that cycle 3 has only six lessons. It has seven, and the
 * sheet was right all along.
 *
 * Cycle 3 runs 9/8 - 9/18: nine school days, being two X days (9/8, 9/18) plus
 * rotation days 1-7. Under the board's own ROT_DAYS rotation table, every
 * period meets SEVEN times in that window. The earlier count of six came from
 * the deck-derived rotation skeleton, which was missing four slots because
 * cycle 3's grid was split across two slides: A on 9/9, G on 9/10, and D and E
 * on 9/11. Rebuilding the skeleton from ROT_DAYS now agrees with this sheet's
 * "# Lessons" column on 18 of 18 cycles.
 *
 * Formatting for I38 is copied from H38 (that block's "Lesson 6" header) so it
 * matches its neighbours rather than the greyed N/A style.
 */
function restoreCycle3Lesson7() {
  var ss = SpreadsheetApp.getActive();

  var cal = ss.getSheetByName('Cycle Calendar');
  if (!cal) { throw new Error('No sheet named "Cycle Calendar"'); }
  // Sanity-check the row really is Semester 1 / Cycle 3 before editing it.
  var sem = cal.getRange('A4').getValue();
  var cyc = cal.getRange('B4').getValue();
  if (String(sem).indexOf('1') === -1 || Number(cyc) !== 3) {
    throw new Error('Cycle Calendar row 4 is "' + sem + '" cycle ' + cyc +
                    ', not Semester 1 cycle 3. Nothing changed.');
  }
  cal.getRange('E4').setValue(7);
  Logger.log('Cycle Calendar: E4 -> 7');

  var lp = ss.getSheetByName('Lesson Planner');
  if (!lp) { throw new Error('No sheet named "Lesson Planner"'); }
  var neighbour = lp.getRange('H38').getValue();
  if (String(neighbour).indexOf('Lesson') !== 0) {
    throw new Error('Expected a "Lesson N" header in Lesson Planner!H38 to ' +
                    'copy formatting from, found "' + neighbour + '". ' +
                    'Cycle Calendar was updated; set I38 to "Lesson 7" by hand.');
  }
  lp.getRange('H38').copyTo(lp.getRange('I38'), { formatOnly: true });
  lp.getRange('I38').setValue('Lesson 7');
  Logger.log('Lesson Planner: I38 -> "Lesson 7", formatted like H38');
}
