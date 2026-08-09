/*******************************************************
 * COLOR WAR 2026
 * SCORE ENTRY + SCORE VIEWER
 *******************************************************/

const SPREADSHEET_ID =
  '1e_2ATVYOoEmbfOPh2YxkN7g_V5_dR_gI-j9JRWG-QLg';

const SHEETS = {
  'Day 1': 'Day 1',
  'Day 2': 'Day 2'
};

const LOG_SHEET = 'Log';


/*******************************************************
 * JUDGE COLUMNS
 *
 * Judge 1  = H:I
 * Judge 2  = K:L
 * Judge 3  = N:O
 * Judge 4  = Q:R
 * Judge 5  = T:U
 * Judge 6  = W:X
 * Judge 7  = Z:AA
 * Judge 8  = AC:AD
 * Judge 9  = AF:AG
 * Judge 10 = AI:AJ
 * Judge 11 = AL:AM
 * Judge 12 = AO:AP
 * Judge 13 = AR:AS
 * Judge 14 = AU:AV
 * Judge 15 = AX:AY
 * Judge 16 = BA:BB
 *
 * BC / BD = IGNORED
 * BE       = TOTAL BLUE JUDGE SCORES
 * BF       = TOTAL RED JUDGE SCORES
 *******************************************************/

const JUDGE_PAIRS = [
  [8, 9],     // H:I
  [11, 12],   // K:L
  [14, 15],   // N:O
  [17, 18],   // Q:R
  [20, 21],   // T:U
  [23, 24],   // W:X
  [26, 27],   // Z:AA
  [29, 30],   // AC:AD
  [32, 33],   // AF:AG
  [35, 36],   // AI:AJ
  [38, 39],   // AL:AM
  [41, 42],   // AO:AP
  [44, 45],   // AR:AS
  [47, 48],   // AU:AV
  [50, 51],   // AX:AY
  [53, 54]    // BA:BB
];


/*******************************************************
 * DO GET
 *******************************************************/

function doGet() {

  return HtmlService
    .createHtmlOutput(getHtml())
    .setTitle('Color War 2026')
    .setXFrameOptionsMode(
      HtmlService.XFrameOptionsMode.ALLOWALL
    );

}


/*******************************************************
 * OPEN SPREADSHEET
 *******************************************************/

function getSpreadsheet() {

  return SpreadsheetApp.openById(
    SPREADSHEET_ID
  );

}


/*******************************************************
 * GET SHEET
 *******************************************************/

function getDaySheet(day) {

  const ss = getSpreadsheet();

  if (!SHEETS[day]) {
    throw new Error('Invalid day: ' + day);
  }

  const sheet =
    ss.getSheetByName(SHEETS[day]);

  if (!sheet) {
    throw new Error(
      'Sheet "' + day + '" was not found.'
    );
  }

  return sheet;

}


/*******************************************************
 * LAST ROW
 *******************************************************/

function getLastRow(day) {

  if (day === 'Day 2') {
    return 25;
  }

  return 24;

}


/*******************************************************
 * GET JUDGES
 *******************************************************/

function getJudges(day) {

  const sheet = getDaySheet(day);

  const judges = [];

  JUDGE_PAIRS.forEach(function(pair, index) {

    const blueColumn = pair[0];
    const redColumn = pair[1];

    const name =
      sheet
        .getRange(1, blueColumn)
        .getDisplayValue()
        .trim();

    if (name !== '') {

      judges.push({

        number: index + 1,

        name: name,

        blueColumn: blueColumn,

        redColumn: redColumn

      });

    }

  });

  return judges;

}


/*******************************************************
 * GET EVENTS
 *******************************************************/

function getEvents(day) {

  const sheet = getDaySheet(day);

  const lastRow = getLastRow(day);

  const events = [];

  const range =
    sheet.getRange(
      3,
      2,
      lastRow - 2,
      2
    );

  const values =
    range.getDisplayValues();


  /*
   * Find merged cells in B:B
   */

  const mergedRanges =
    sheet
      .getRange(
        3,
        2,
        lastRow - 2,
        1
      )
      .getMergedRanges();


  const mergedRows = {};


  mergedRanges.forEach(function(r) {

    const start = r.getRow();

    const end =
      start +
      r.getNumRows() -
      1;

    for (
      let row = start;
      row <= end;
      row++
    ) {

      mergedRows[row] = {
        start: start,
        end: end
      };

    }

  });


  for (
    let row = 3;
    row <= lastRow;
    row++
  ) {

    const b =
      sheet
        .getRange(row, 2)
        .getDisplayValue()
        .trim();


    /*
     * MERGED B
     */

    if (mergedRows[row]) {

      const group =
        mergedRows[row];


      if (row !== group.start) {
        continue;
      }


      const groupName =
        sheet
          .getRange(
            group.start,
            2
          )
          .getDisplayValue()
          .trim();


      const options = [];


      for (
        let r = group.start;
        r <= group.end;
        r++
      ) {

        const optionName =
          sheet
            .getRange(r, 3)
            .getDisplayValue()
            .trim();


        options.push({

          row: r,

          name:
            optionName ||
            'Option ' +
            (r - group.start + 1)

        });

      }


      events.push({

        type: 'group',

        name: groupName,

        row: group.start,

        startRow: group.start,

        endRow: group.end,

        options: options

      });


      continue;

    }


    /*
     * NORMAL B
     */

    if (b !== '') {

      events.push({

        type: 'normal',

        name: b,

        row: row,

        options: []

      });

    }

  }

  return events;

}


/*******************************************************
 * GET DAY DATA
 *******************************************************/

function getDayData(day) {

  return {

    events: getEvents(day),

    judges: getJudges(day)

  };

}


/*******************************************************
 * GET EXISTING SCORE
 *******************************************************/

function getScores(
  day,
  row,
  judgeName
) {

  const sheet =
    getDaySheet(day);

  const judges =
    getJudges(day);


  const judge =
    judges.find(function(j) {

      return j.name === judgeName;

    });


  if (!judge) {

    throw new Error(
      'Judge "' +
      judgeName +
      '" was not found.'
    );

  }


  return {

    blue:
      sheet
        .getRange(
          Number(row),
          judge.blueColumn
        )
        .getDisplayValue(),

    red:
      sheet
        .getRange(
          Number(row),
          judge.redColumn
        )
        .getDisplayValue()

  };

}


/*******************************************************
 * SAVE SCORE
 *******************************************************/

function saveScores(
  day,
  row,
  eventName,
  optionName,
  judgeName,
  blue,
  red
) {

  const sheet =
    getDaySheet(day);

  const judges =
    getJudges(day);


  const judge =
    judges.find(function(j) {

      return j.name === judgeName;

    });


  if (!judge) {

    throw new Error(
      'Judge not found.'
    );

  }


  const rowNumber =
    Number(row);


  const oldBlue =
    sheet
      .getRange(
        rowNumber,
        judge.blueColumn
      )
      .getDisplayValue();


  const oldRed =
    sheet
      .getRange(
        rowNumber,
        judge.redColumn
      )
      .getDisplayValue();


  const action =
    oldBlue !== '' ||
    oldRed !== ''
      ? 'Updated'
      : 'New Entry';


  /*
   * BLUE
   */

  if (
    blue !== '' &&
    blue !== null &&
    blue !== undefined
  ) {

    if (isNaN(Number(blue))) {

      throw new Error(
        'Blue score must be a number.'
      );

    }

    sheet
      .getRange(
        rowNumber,
        judge.blueColumn
      )
      .setValue(
        Number(blue)
      );

  } else {

    sheet
      .getRange(
        rowNumber,
        judge.blueColumn
      )
      .clearContent();

  }


  /*
   * RED
   */

  if (
    red !== '' &&
    red !== null &&
    red !== undefined
  ) {

    if (isNaN(Number(red))) {

      throw new Error(
        'Red score must be a number.'
      );

    }

    sheet
      .getRange(
        rowNumber,
        judge.redColumn
      )
      .setValue(
        Number(red)
      );

  } else {

    sheet
      .getRange(
        rowNumber,
        judge.redColumn
      )
      .clearContent();

  }


  SpreadsheetApp.flush();


  /*
   * LOG
   */

  const ss =
    getSpreadsheet();

  const log =
    ss.getSheetByName(
      LOG_SHEET
    );


  if (log) {

    log.appendRow([

      new Date(),

      day,

      eventName,

      optionName,

      judgeName,

      blue,

      red,

      action

    ]);

  }


  return true;

}


/*******************************************************
 * GET ROW SCORE DATA
 *******************************************************/

function getRowScoreData(
  sheet,
  row,
  judges
) {

  /*
   * E = final Blue
   * F = final Red
   * G = number of judges
   * BE = raw Blue total
   * BF = raw Red total
   */

  const finalBlue =
    sheet
      .getRange(row, 5)
      .getDisplayValue()
      .trim();


  const finalRed =
    sheet
      .getRange(row, 6)
      .getDisplayValue()
      .trim();


  const judgeCount =
    sheet
      .getRange(row, 7)
      .getDisplayValue()
      .trim();


  /*
   * BE = column 57
   * BF = column 58
   */

  const rawBlue =
    sheet
      .getRange(row, 57)
      .getDisplayValue()
      .trim();


  const rawRed =
    sheet
      .getRange(row, 58)
      .getDisplayValue()
      .trim();


  const judgeScores = [];


  judges.forEach(function(judge) {

    const blue =
      sheet
        .getRange(
          row,
          judge.blueColumn
        )
        .getDisplayValue()
        .trim();


    const red =
      sheet
        .getRange(
          row,
          judge.redColumn
        )
        .getDisplayValue()
        .trim();


    if (
      blue !== '' ||
      red !== ''
    ) {

      judgeScores.push({

        name: judge.name,

        blue: blue,

        red: red

      });

    }

  });


  return {

    finalBlue: finalBlue,

    finalRed: finalRed,

    judgeCount: judgeCount,

    rawBlue: rawBlue,

    rawRed: rawRed,

    judges: judgeScores

  };

}


/*******************************************************
 * GET OVERALL TOTALS
 *******************************************************/

function getOverallTotals(day) {

  const sheet =
    getDaySheet(day);


  let blueCell;
  let redCell;


  if (day === 'Day 1') {

    blueCell = 'E26';
    redCell = 'F26';

  } else {

    blueCell = 'E28';
    redCell = 'F28';

  }


  return {

    blue:
      sheet
        .getRange(blueCell)
        .getDisplayValue()
        .trim(),

    red:
      sheet
        .getRange(redCell)
        .getDisplayValue()
        .trim()

  };

}


/*******************************************************
 * GET BOTH DAY TOTALS
 *******************************************************/

function getAllTotals() {

  return {

    day1:
      getOverallTotals('Day 1'),

    day2:
      getOverallTotals('Day 2')

  };

}


/*******************************************************
 * GET SCORE VIEW
 *******************************************************/

function getScoreView(day) {

  const sheet =
    getDaySheet(day);

  const judges =
    getJudges(day);

  const events =
    getEvents(day);

  const totals =
    getAllTotals();


  const report = [];


  events.forEach(function(event) {


    /*
     * NORMAL ACTIVITY
     */

    if (event.type === 'normal') {

      const scores =
        getRowScoreData(
          sheet,
          event.row,
          judges
        );


      /*
       * Only show activities
       * where there is actually
       * a score.
       */

      if (
        scores.judges.length > 0 ||
        scores.finalBlue !== '' ||
        scores.finalRed !== ''
      ) {

        report.push({

          type: 'normal',

          name: event.name,

          row: event.row,

          scores: scores

        });

      }

      return;

    }


    /*
     * MERGED GROUP
     */

    if (event.type === 'group') {

      const rows = [];


      event.options.forEach(function(option) {

        const scores =
          getRowScoreData(
            sheet,
            option.row,
            judges
          );


        /*
         * Only include rows
         * that actually have scores.
         */

        if (
          scores.judges.length > 0 ||
          scores.finalBlue !== '' ||
          scores.finalRed !== ''
        ) {

          rows.push({

            row: option.row,

            name: option.name,

            scores: scores

          });

        }

      });


      if (rows.length > 0) {

        report.push({

          type: 'group',

          name: event.name,

          rows: rows

        });

      }

    }

  });


  return {

    day: day,

    totals: totals,

    report: report

  };

}


/*******************************************************
 * HTML
 *******************************************************/

function getHtml() {

return `

<!DOCTYPE html>

<html>

<head>

<meta name="viewport"
content="width=device-width, initial-scale=1">

<title>Color War 2026</title>


<style>

/*******************************************************
 * GENERAL
 *******************************************************/

* {
  box-sizing: border-box;
}


body {

  margin: 0;

  font-family:
    Arial,
    Helvetica,
    sans-serif;

  background-color:
    #f1f3f4;

  background-image:
    url('data:image/png;base64,${BACKGROUND_IMAGE_BASE64}');

  background-size:
    cover;

  background-position:
    top center;

  background-repeat:
    no-repeat;

  background-attachment:
    fixed;

  color:
    #202124;

}


button,
input,
select {
  font-family: inherit;
}


/*******************************************************
 * HEADER
 *******************************************************/

.header {

  position: sticky;

  top: 0;

  z-index: 1000;

  background:
    white;

  box-shadow:
    0 2px 10px
    rgba(0,0,0,.15);

}


.brand {

  text-align:
    center;

  font-size:
    28px;

  font-weight:
    900;

  padding:
    15px 10px 10px;

}


.toggle {

  max-width:
    700px;

  margin:
    0 auto 8px;

  padding:
    4px;

  background:
    #e8eaed;

  border-radius:
    12px;

  display:
    flex;

}


.toggle button {

  flex:
    1;

  border:
    none;

  background:
    transparent;

  padding:
    11px;

  border-radius:
    9px;

  font-weight:
    800;

  cursor:
    pointer;

}


.toggle button.active {

  background:
    #202124;

  color:
    white;

}


.dayToggle button.active {

  background:
    #1a73e8;

}


/*******************************************************
 * CONTAINER
 *******************************************************/

.container {

  max-width:
    1200px;

  margin:
    auto;

  padding:
    18px;

}


.page {
  display: none;
}


.page.active {
  display: block;
}


/*******************************************************
 * EVENT BUTTONS
 *******************************************************/

.title {

  text-align:
    center;

  font-size:
    24px;

  font-weight:
    900;

  margin:
    4px 0 18px;

}


.eventGrid {

  display:
    grid;

  grid-template-columns:
    repeat(2, 1fr);

  gap:
    12px;

}


.eventButton {

  border:
    none;

  background:
    white;

  border-radius:
    14px;

  padding:
    23px 12px;

  font-size:
    17px;

  font-weight:
    800;

  cursor:
    pointer;

  box-shadow:
    0 2px 8px
    rgba(0,0,0,.12);

}


/*******************************************************
 * SCORE ENTRY
 *******************************************************/

.backButton {

  width:
    100%;

  border:
    none;

  border-radius:
    9px;

  padding:
    11px;

  margin-bottom:
    14px;

  background:
    #5f6368;

  color:
    white;

  font-weight:
    800;

}


.eventTitle {

  text-align:
    center;

  font-size:
    25px;

  font-weight:
    900;

  margin-bottom:
    14px;

}


.optionGrid {

  display:
    grid;

  grid-template-columns:
    repeat(3, 1fr);

  gap:
    8px;

  margin-bottom:
    15px;

}


.optionButton {

  border:
    2px solid
    #dadce0;

  background:
    white;

  border-radius:
    9px;

  padding:
    12px 8px;

  font-weight:
    800;

}


.optionButton.active {

  background:
    #1a73e8;

  border-color:
    #1a73e8;

  color:
    white;

}


.judgeSelect {

  width:
    100%;

  padding:
    13px;

  border-radius:
    9px;

  border:
    1px solid
    #ccc;

  background:
    white;

  font-size:
    17px;

  margin-bottom:
    14px;

}


.scoreGrid {

  display:
    grid;

  grid-template-columns:
    1fr 1fr;

  gap:
    12px;

}


.scoreBox {

  border-radius:
    14px;

  padding:
    15px;

  text-align:
    center;

}


.blueBox {
  background:
    #e8f0fe;
}


.redBox {
  background:
    #fce8e6;
}


.scoreLabel {

  font-size:
    18px;

  font-weight:
    900;

  margin-bottom:
    8px;

}


.scoreInput {

  width:
    100%;

  padding:
    13px;

  font-size:
    25px;

  text-align:
    center;

  border:
    1px solid
    #ccc;

  border-radius:
    8px;

}


.saveButton {

  width:
    100%;

  padding:
    15px;

  margin-top:
    13px;

  border:
    none;

  border-radius:
    9px;

  background:
    #188038;

  color:
    white;

  font-size:
    18px;

  font-weight:
    900;

}


.message {

  text-align:
    center;

  font-weight:
    900;

  padding:
    10px;

}


/*******************************************************
 * TOTALS
 *******************************************************/

.totalTitle {

  text-align:
    center;

  font-size:
    15px;

  font-weight:
    900;

  color:
    #5f6368;

  margin-bottom:
    6px;

}


.totalGrid {

  display:
    grid;

  grid-template-columns:
    1fr 1fr;

  gap:
    10px;

}


.totalBox {

  color:
    white;

  border-radius:
    15px;

  padding:
    14px;

  text-align:
    center;

  box-shadow:
    0 3px 9px
    rgba(0,0,0,.17);

}


.blueTotal {
  background:
    #1a73e8;
}


.redTotal {
  background:
    #d93025;
}


.totalTeam {

  font-size:
    14px;

  font-weight:
    900;

}


.totalNumber {

  font-size:
    31px;

  font-weight:
    900;

}


.dayTotals {

  display:
    grid;

  grid-template-columns:
    1fr 1fr;

  gap:
    8px;

  margin:
    10px 0 20px;

}


.dayTotal {

  background:
    white;

  border-radius:
    10px;

  padding:
    9px;

  text-align:
    center;

  box-shadow:
    0 1px 5px
    rgba(0,0,0,.10);

}


.dayName {

  font-size:
    11px;

  font-weight:
    900;

  color:
    #666;

}


.dayNumbers {

  font-size:
    15px;

  font-weight:
    900;

}


.dayBlue {
  color:
    #1a73e8;
}


.dayRed {
  color:
    #d93025;
}


/*******************************************************
 * SCORE VIEW
 *******************************************************/

.activity {

  background:
    white;

  border-radius:
    15px;

  padding:
    12px;

  margin-bottom:
    12px;

  box-shadow:
    0 2px 8px
    rgba(0,0,0,.11);

}


.activityName {

  font-size:
    18px;

  font-weight:
    900;

  margin-bottom:
    7px;

}


.summary {

  display:
    flex;

  flex-wrap:
    wrap;

  gap:
    5px;

  margin-bottom:
    7px;

}


.summaryBox {

  border-radius:
    7px;

  padding:
    4px 7px;

  font-size:
    11px;

  font-weight:
    900;

}


.finalBlue {

  background:
    #e8f0fe;

  color:
    #174ea6;

}


.finalRed {

  background:
    #fce8e6;

  color:
    #a50e0e;

}


.judgeCount {

  background:
    #f1f3f4;

}


.rawTotals {

  display:
    flex;

  flex-wrap:
    wrap;

  gap:
    5px;

  margin-bottom:
    8px;

}


.rawBlue {

  background:
    #d2e3fc;

  color:
    #174ea6;

}


.rawRed {

  background:
    #fad2cf;

  color:
    #a50e0e;

}


/*******************************************************
 * GROUP BUBBLE
 *******************************************************/

.group {

  background:
    white;

  border:
    3px solid
    #dadce0;

  border-radius:
    23px;

  padding:
    14px;

  margin-bottom:
    16px;

  box-shadow:
    0 3px 10px
    rgba(0,0,0,.13);

}


.groupName {

  text-align:
    center;

  font-size:
    21px;

  font-weight:
    900;

  padding-bottom:
    9px;

  margin-bottom:
    10px;

  border-bottom:
    2px solid
    #eee;

}


.groupOption {

  background:
    #f8f9fa;

  border-radius:
    14px;

  padding:
    9px;

  margin-bottom:
    9px;

}


.optionName {

  font-size:
    15px;

  font-weight:
    900;

  margin-bottom:
    6px;

}


/*******************************************************
 * JUDGE BUBBLES
 *******************************************************/

.judgeContainer {

  display:
    flex;

  flex-wrap:
    wrap;

  gap:
    7px;

}


.judgeBubble {

  flex:
    1 1 175px;

  min-width:
    160px;

  max-width:
    260px;

  border:
    2px solid
    #dadce0;

  border-radius:
    999px;

  background:
    #fafafa;

  padding:
    8px 11px;

  text-align:
    center;

  box-shadow:
    0 2px 5px
    rgba(0,0,0,.09);

}


.judgeName {

  font-size:
    11px;

  font-weight:
    900;

  white-space:
    nowrap;

  overflow:
    hidden;

  text-overflow:
    ellipsis;

  margin-bottom:
    4px;

}


.judgeScoreLine {

  display:
    flex;

  justify-content:
    center;

  gap:
    5px;

}


.judgeBlue {

  background:
    #4285f4;

  color:
    white;

  border-radius:
    30px;

  padding:
    4px 8px;

  font-size:
    12px;

  font-weight:
    900;

}


.judgeRed {

  background:
    #ea4335;

  color:
    white;

  border-radius:
    30px;

  padding:
    4px 8px;

  font-size:
    12px;

  font-weight:
    900;

}


.empty {

  background:
    white;

  border-radius:
    14px;

  padding:
    35px;

  text-align:
    center;

  color:
    #666;

}


/*******************************************************
 * MOBILE
 *******************************************************/

@media(max-width:700px) {

  .brand {
    font-size: 23px;
  }

  .container {
    padding: 10px;
  }

  .eventGrid {
    grid-template-columns:
      1fr;
  }

  .optionGrid {
    grid-template-columns:
      repeat(2, 1fr);
  }

  .scoreGrid {
    grid-template-columns:
      1fr;
  }

  .judgeBubble {
    min-width:
      145px;

    flex-basis:
      145px;
  }

}

</style>

</head>


<body>


<!-- =====================================================
     HEADER
===================================================== -->

<div class="header">

  <div class="brand">
    COLOR WAR 2026
  </div>


  <div class="toggle">

    <button
      id="enterButton"
      class="active"
      onclick="setMode('enter')">

      ENTER SCORES

    </button>


    <button
      id="viewButton"
      onclick="setMode('view')">

      VIEW SCORES

    </button>

  </div>


  <div class="toggle dayToggle">

    <button
      id="day1Button"
      class="active"
      onclick="changeDay('Day 1')">

      DAY 1

    </button>


    <button
      id="day2Button"
      onclick="changeDay('Day 2')">

      DAY 2

    </button>

  </div>

</div>


<!-- =====================================================
     CONTENT
===================================================== -->

<div class="container">


<!-- =====================================================
     ENTER PAGE
===================================================== -->

<div
  id="enterPage"
  class="page active">


  <div
    id="eventList">


    <div class="title">
      SELECT ACTIVITY
    </div>


    <div
      id="eventGrid"
      class="eventGrid">

      Loading...

    </div>

  </div>


  <div
    id="entryPage"
    style="display:none;">


    <button
      class="backButton"
      onclick="backToEvents()">

      ← BACK TO ACTIVITIES

    </button>


    <div
      id="eventTitle"
      class="eventTitle">
    </div>


    <div
      id="optionGrid"
      class="optionGrid">
    </div>


    <select
      id="judgeSelect"
      class="judgeSelect"
      onchange="loadExistingScore()">

      <option value="">
        SELECT JUDGE
      </option>

    </select>


    <div class="scoreGrid">


      <div class="scoreBox blueBox">

        <div class="scoreLabel">
          🔵 BLUE SCORE
        </div>

        <input
          id="blueScore"
          class="scoreInput"
          type="number"
          step="any">

      </div>


      <div class="scoreBox redBox">

        <div class="scoreLabel">
          🔴 RED SCORE
        </div>

        <input
          id="redScore"
          class="scoreInput"
          type="number"
          step="any">

      </div>


    </div>


    <button
      id="saveButton"
      class="saveButton"
      onclick="saveScore()">

      SAVE SCORE

    </button>


    <div
      id="message"
      class="message">
    </div>


  </div>

</div>


<!-- =====================================================
     VIEW PAGE
===================================================== -->

<div
  id="viewPage"
  class="page">


  <div class="title">
    SCORE RESULTS
  </div>


  <div
    id="scoreView">

    Loading...

  </div>


</div>


</div>


<script>

/*******************************************************
 * VARIABLES
 *******************************************************/

let currentDay = 'Day 1';

let currentMode = 'enter';

let dayData = {};

let currentEvent = null;

let currentOption = null;


/*******************************************************
 * MODE
 *******************************************************/

function setMode(mode) {

  currentMode = mode;


  document
    .getElementById('enterButton')
    .classList
    .toggle(
      'active',
      mode === 'enter'
    );


  document
    .getElementById('viewButton')
    .classList
    .toggle(
      'active',
      mode === 'view'
    );


  document
    .getElementById('enterPage')
    .classList
    .toggle(
      'active',
      mode === 'enter'
    );


  document
    .getElementById('viewPage')
    .classList
    .toggle(
      'active',
      mode === 'view'
    );


  if (mode === 'view') {

    loadScoreView();

  }

}


/*******************************************************
 * CHANGE DAY
 *******************************************************/

function changeDay(day) {

  currentDay = day;


  document
    .getElementById('day1Button')
    .classList
    .toggle(
      'active',
      day === 'Day 1'
    );


  document
    .getElementById('day2Button')
    .classList
    .toggle(
      'active',
      day === 'Day 2'
    );


  currentEvent = null;

  currentOption = null;


  if (currentMode === 'enter') {

    document
      .getElementById('eventList')
      .style.display =
      'block';


    document
      .getElementById('entryPage')
      .style.display =
      'none';


    loadDay();

  } else {

    loadScoreView();

  }

}


/*******************************************************
 * LOAD DAY
 *******************************************************/

function loadDay() {

  document
    .getElementById('eventGrid')
    .innerHTML =
    '<div class="empty">Loading...</div>';


  google.script.run

    .withSuccessHandler(function(data) {

      dayData = data;

      renderEvents();

    })

    .withFailureHandler(function(error) {

      document
        .getElementById('eventGrid')
        .innerHTML =
        '<div class="empty">' +
        error.message +
        '</div>';

    })

    .getDayData(currentDay);

}


/*******************************************************
 * RENDER EVENTS
 *******************************************************/

function renderEvents() {

  const grid =
    document.getElementById(
      'eventGrid'
    );


  grid.innerHTML = '';


  if (
    !dayData.events ||
    dayData.events.length === 0
  ) {

    grid.innerHTML =
      '<div class="empty">' +
      'No activities found.' +
      '</div>';

    return;

  }


  dayData.events.forEach(function(event) {

    const button =
      document.createElement(
        'button'
      );


    button.className =
      'eventButton';


    button.textContent =
      event.name;


    button.onclick =
      function() {

        openEvent(event);

      };


    grid.appendChild(button);

  });

}


/*******************************************************
 * OPEN EVENT
 *******************************************************/

function openEvent(event) {

  currentEvent = event;

  currentOption = null;


  document
    .getElementById('eventList')
    .style.display =
    'none';


  document
    .getElementById('entryPage')
    .style.display =
    'block';


  document
    .getElementById('eventTitle')
    .textContent =
    event.name;


  renderOptions();

  loadJudges();


  if (event.type === 'normal') {

    currentOption = {

      row: event.row,

      name: ''

    };

  }

}


/*******************************************************
 * OPTIONS
 *******************************************************/

function renderOptions() {

  const grid =
    document.getElementById(
      'optionGrid'
    );


  grid.innerHTML = '';


  if (
    !currentEvent ||
    currentEvent.type !== 'group'
  ) {

    return;

  }


  currentEvent.options.forEach(
    function(option) {

      const button =
        document.createElement(
          'button'
        );


      button.className =
        'optionButton';


      button.textContent =
        option.name;


      button.dataset.row =
        option.row;


      button.onclick =
        function() {

          selectOption(option);

        };


      grid.appendChild(button);

    }
  );


  if (
    currentEvent.options.length > 0
  ) {

    selectOption(
      currentEvent.options[0]
    );

  }

}


/*******************************************************
 * SELECT OPTION
 *******************************************************/

function selectOption(option) {

  currentOption = option;


  document
    .querySelectorAll(
      '.optionButton'
    )
    .forEach(function(button) {

      button.classList.toggle(

        'active',

        Number(
          button.dataset.row
        ) ===
        Number(option.row)

      );

    });


  const judge =
    document
      .getElementById(
        'judgeSelect'
      )
      .value;


  if (judge) {

    loadExistingScore();

  } else {

    clearScores();

  }

}


/*******************************************************
 * LOAD JUDGES
 *******************************************************/

function loadJudges() {

  const select =
    document.getElementById(
      'judgeSelect'
    );


  select.innerHTML =
    '<option value="">SELECT JUDGE</option>';


  dayData.judges.forEach(
    function(judge) {

      const option =
        document.createElement(
          'option'
        );


      option.value =
        judge.name;


      option.textContent =
        judge.name;


      select.appendChild(option);

    }
  );


  clearScores();

}


/*******************************************************
 * CLEAR SCORES
 *******************************************************/

function clearScores() {

  document
    .getElementById('blueScore')
    .value = '';


  document
    .getElementById('redScore')
    .value = '';


  document
    .getElementById('message')
    .textContent = '';

}


/*******************************************************
 * LOAD EXISTING SCORE
 *******************************************************/

function loadExistingScore() {

  const judge =
    document
      .getElementById(
        'judgeSelect'
      )
      .value;


  if (
    !judge ||
    !currentOption
  ) {

    clearScores();

    return;

  }


  google.script.run

    .withSuccessHandler(
      function(result) {

        document
          .getElementById(
            'blueScore'
          )
          .value =
          result.blue;


        document
          .getElementById(
            'redScore'
          )
          .value =
          result.red;

      }
    )

    .withFailureHandler(
      function(error) {

        alert(
          error.message
        );

      }
    )

    .getScores(

      currentDay,

      currentOption.row,

      judge

    );

}


/*******************************************************
 * SAVE SCORE
 *******************************************************/

function saveScore() {

  const judge =
    document
      .getElementById(
        'judgeSelect'
      )
      .value;


  const blue =
    document
      .getElementById(
        'blueScore'
      )
      .value;


  const red =
    document
      .getElementById(
        'redScore'
      )
      .value;


  if (!currentOption) {

    alert(
      'Please select an activity.'
    );

    return;

  }


  if (!judge) {

    alert(
      'Please select a judge.'
    );

    return;

  }


  if (
    blue === '' &&
    red === ''
  ) {

    alert(
      'Please enter a score.'
    );

    return;

  }


  const button =
    document.getElementById(
      'saveButton'
    );


  button.disabled = true;

  button.textContent =
    'SAVING...';


  google.script.run

    .withSuccessHandler(
      function() {

        button.disabled = false;

        button.textContent =
          'SAVE SCORE';


        document
          .getElementById(
            'message'
          )
          .textContent =
          '✓ SCORE SAVED';


        /*
         * Automatically move to
         * the next option in a
         * merged group.
         */

        if (
          currentEvent &&
          currentEvent.type === 'group'
        ) {

          const index =
            currentEvent.options
              .findIndex(
                function(o) {

                  return Number(o.row) ===
                    Number(currentOption.row);

                }
              );


          if (
            index >= 0 &&
            index <
            currentEvent.options.length - 1
          ) {

            selectOption(
              currentEvent.options[index + 1]
            );

          }

        }

      }
    )

    .withFailureHandler(
      function(error) {

        button.disabled = false;

        button.textContent =
          'SAVE SCORE';


        document
          .getElementById(
            'message'
          )
          .textContent =
          'ERROR: ' +
          error.message;

      }
    )

    .saveScores(

      currentDay,

      currentOption.row,

      currentEvent.name,

      currentOption.name,

      judge,

      blue,

      red

    );

}


/*******************************************************
 * BACK TO EVENTS
 *******************************************************/

function backToEvents() {

  currentEvent = null;

  currentOption = null;


  document
    .getElementById('entryPage')
    .style.display =
    'none';


  document
    .getElementById('eventList')
    .style.display =
    'block';


  loadDay();

}


/*******************************************************
 * LOAD SCORE VIEW
 *******************************************************/

function loadScoreView() {

  document
    .getElementById('scoreView')
    .innerHTML =
    '<div class="empty">Loading scores...</div>';


  google.script.run

    .withSuccessHandler(
      function(data) {

        renderScoreView(data);

      }
    )

    .withFailureHandler(
      function(error) {

        document
          .getElementById(
            'scoreView'
          )
          .innerHTML =
          '<div class="empty">' +
          error.message +
          '</div>';

      }
    )

    .getScoreView(currentDay);

}


/*******************************************************
 * RENDER SCORE VIEW
 *******************************************************/

function renderScoreView(data) {

  const container =
    document.getElementById(
      'scoreView'
    );


  container.innerHTML = '';


  /*
   * TOTAL TITLE
   */

  const totalTitle =
    document.createElement(
      'div'
    );


  totalTitle.className =
    'totalTitle';


  totalTitle.textContent =
    'COLOR WAR TOTAL';


  container.appendChild(
    totalTitle
  );


  /*
   * COMBINED TOTALS
   *
   * Uses the existing Day 1
   * and Day 2 total cells.
   */

  const combinedBlue =
    numericValue(
      data.totals.day1.blue
    ) +
    numericValue(
      data.totals.day2.blue
    );


  const combinedRed =
    numericValue(
      data.totals.day1.red
    ) +
    numericValue(
      data.totals.day2.red
    );


  const totalGrid =
    document.createElement(
      'div'
    );


  totalGrid.className =
    'totalGrid';


  totalGrid.innerHTML =

    '<div class="totalBox blueTotal">' +

      '<div class="totalTeam">' +
        '🔵 BLUE' +
      '</div>' +

      '<div class="totalNumber">' +
        formatNumber(combinedBlue) +
      '</div>' +

    '</div>' +


    '<div class="totalBox redTotal">' +

      '<div class="totalTeam">' +
        '🔴 RED' +
      '</div>' +

      '<div class="totalNumber">' +
        formatNumber(combinedRed) +
      '</div>' +

    '</div>';


  container.appendChild(
    totalGrid
  );


  /*
   * DAY TOTALS
   */

  const dayTotals =
    document.createElement(
      'div'
    );


  dayTotals.className =
    'dayTotals';


  dayTotals.innerHTML =

    '<div class="dayTotal">' +

      '<div class="dayName">' +
        'DAY 1' +
      '</div>' +

      '<div class="dayNumbers">' +

        '<span class="dayBlue">' +
          '🔵 ' +
          escapeHtml(
            data.totals.day1.blue
          ) +
        '</span>' +

        ' &nbsp; ' +

        '<span class="dayRed">' +
          '🔴 ' +
          escapeHtml(
            data.totals.day1.red
          ) +
        '</span>' +

      '</div>' +

    '</div>' +


    '<div class="dayTotal">' +

      '<div class="dayName">' +
        'DAY 2' +
      '</div>' +

      '<div class="dayNumbers">' +

        '<span class="dayBlue">' +
          '🔵 ' +
          escapeHtml(
            data.totals.day2.blue
          ) +
        '</span>' +

        ' &nbsp; ' +

        '<span class="dayRed">' +
          '🔴 ' +
          escapeHtml(
            data.totals.day2.red
          ) +
        '</span>' +

      '</div>' +

    '</div>';


  container.appendChild(
    dayTotals
  );


  /*
   * NO SCORES
   */

  if (
    !data.report ||
    data.report.length === 0
  ) {

    const empty =
      document.createElement(
        'div'
      );


    empty.className =
      'empty';


    empty.textContent =
      'No scores have been entered for ' +
      data.day +
      '.';


    container.appendChild(
      empty
    );


    return;

  }


  /*
   * ACTIVITIES
   */

  data.report.forEach(
    function(item) {


      /*
       * NORMAL
       */

      if (
        item.type === 'normal'
      ) {

        const activity =
          document.createElement(
            'div'
          );


        activity.className =
          'activity';


        activity.appendChild(
          createActivityContent(
            item.name,
            item.scores
          )
        );


        container.appendChild(
          activity
        );


        return;

      }


      /*
       * MERGED GROUP
       */

      if (
        item.type === 'group'
      ) {

        const group =
          document.createElement(
            'div'
          );


        group.className =
          'group';


        const groupName =
          document.createElement(
            'div'
          );


        groupName.className =
          'groupName';


        groupName.textContent =
          item.name;


        group.appendChild(
          groupName
        );


        item.rows.forEach(
          function(row) {

            const option =
              document.createElement(
                'div'
              );


            option.className =
              'groupOption';


            const optionName =
              document.createElement(
                'div'
              );


            optionName.className =
              'optionName';


            optionName.textContent =
              row.name;


            option.appendChild(
              optionName
            );


            option.appendChild(
              createSummary(
                row.scores
              )
            );


            option.appendChild(
              createRawTotals(
                row.scores
              )
            );


            option.appendChild(
              createJudgeBubbles(
                row.scores.judges
              )
            );


            group.appendChild(
              option
            );

          }
        );


        container.appendChild(
          group
        );

      }

    }
  );

}


/*******************************************************
 * CREATE NORMAL ACTIVITY
 *******************************************************/

function createActivityContent(
  name,
  scores
) {

  const fragment =
    document.createDocumentFragment();


  const title =
    document.createElement(
      'div'
    );


  title.className =
    'activityName';


  title.textContent =
    name;


  fragment.appendChild(
    title
  );


  fragment.appendChild(
    createSummary(
      scores
    )
  );


  fragment.appendChild(
    createRawTotals(
      scores
    )
  );


  fragment.appendChild(
    createJudgeBubbles(
      scores.judges
    )
  );


  return fragment;

}


/*******************************************************
 * SUMMARY
 *******************************************************/

function createSummary(
  scores
) {

  const summary =
    document.createElement(
      'div'
    );


  summary.className =
    'summary';


  const blue =
    document.createElement(
      'div'
    );


  blue.className =
    'summaryBox finalBlue';


  blue.textContent =
    '🔵 Final Blue: ' +
    scores.finalBlue;


  const red =
    document.createElement(
      'div'
    );


  red.className =
    'summaryBox finalRed';


  red.textContent =
    '🔴 Final Red: ' +
    scores.finalRed;


  const count =
    document.createElement(
      'div'
    );


  count.className =
    'summaryBox judgeCount';


  count.textContent =
    '👥 Judges: ' +
    scores.judgeCount;


  summary.appendChild(
    blue
  );


  summary.appendChild(
    red
  );


  summary.appendChild(
    count
  );


  return summary;

}


/*******************************************************
 * RAW TOTALS
 *******************************************************/

function createRawTotals(
  scores
) {

  const container =
    document.createElement(
      'div'
    );


  container.className =
    'rawTotals';


  /*
   * THESE ARE READ DIRECTLY
   * FROM BE AND BF.
   *
   * NO CALCULATION HERE.
   */

  const blue =
    document.createElement(
      'div'
    );


  blue.className =
    'summaryBox rawBlue';


  blue.textContent =
    'All Judge Blue: ' +
    scores.rawBlue;


  const red =
    document.createElement(
      'div'
    );


  red.className =
    'summaryBox rawRed';


  red.textContent =
    'All Judge Red: ' +
    scores.rawRed;


  container.appendChild(
    blue
  );


  container.appendChild(
    red
  );


  return container;

}


/*******************************************************
 * JUDGE BUBBLES
 *******************************************************/

function createJudgeBubbles(
  judges
) {

  const container =
    document.createElement(
      'div'
    );


  container.className =
    'judgeContainer';


  if (
    !judges ||
    judges.length === 0
  ) {

    return container;

  }


  judges.forEach(
    function(judge) {

      const bubble =
        document.createElement(
          'div'
        );


      bubble.className =
        'judgeBubble';


      const name =
        document.createElement(
          'div'
        );


      name.className =
        'judgeName';


      name.textContent =
        judge.name;


      bubble.appendChild(
        name
      );


      const line =
        document.createElement(
          'div'
        );


      line.className =
        'judgeScoreLine';


      if (
        judge.blue !== ''
      ) {

        const blue =
          document.createElement(
            'span'
          );


        blue.className =
          'judgeBlue';


        blue.textContent =
          '🔵 ' +
          judge.blue;


        line.appendChild(
          blue
        );

      }


      if (
        judge.red !== ''
      ) {

        const red =
          document.createElement(
            'span'
          );


        red.className =
          'judgeRed';


        red.textContent =
          '🔴 ' +
          judge.red;


        line.appendChild(
          red
        );

      }


      bubble.appendChild(
        line
      );


      container.appendChild(
        bubble
      );

    }
  );


  return container;

}


/*******************************************************
 * NUMBER
 *******************************************************/

function numericValue(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {

    return 0;

  }


  const n =
    Number(
      String(value)
        .replace(/,/g, '')
    );


  return isNaN(n)
    ? 0
    : n;

}


/*******************************************************
 * FORMAT NUMBER
 *******************************************************/

function formatNumber(value) {

  if (
    value === '' ||
    value === null ||
    value === undefined
  ) {

    return '';

  }


  const n =
    Number(value);


  if (isNaN(n)) {
    return value;
  }


  if (
    Number.isInteger(n)
  ) {

    return String(n);

  }


  return n
    .toFixed(2)
    .replace(/0+$/, '')
    .replace(/\.$/, '');

}


/*******************************************************
 * ESCAPE HTML
 *******************************************************/

function escapeHtml(value) {

  if (
    value === null ||
    value === undefined
  ) {

    return '';

  }


  return String(value)

    .replace(
      /&/g,
      '&amp;'
    )

    .replace(
      /</g,
      '&lt;'
    )

    .replace(
      />/g,
      '&gt;'
    )

    .replace(
      /"/g,
      '&quot;'
    )

    .replace(
      /'/g,
      '&#039;'
    );

}


/*******************************************************
 * START
 *******************************************************/

loadDay();

</script>

</body>

</html>

`;

}
