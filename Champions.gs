/* -------------------------------------------------------------------------------------
/
/ This is the client-end GoogleApp Script that resides within the CHAMPION Parkrunners
/ Google Spreadsheet.  It complements the server side JavaScript index.js
/ It primarily consist of a series of five macro-based functions aimed at automating the
/ capture and presentation of results for parkrunners related to CHAMPION family.
/ The primary six entry-point functions that are bound to macros/keys are:
/   1.  Re-protect results for each runner
/       (ReprotectEachRunnerResultsSheets - Ctrl+Alt+Shift+7)
/   2.  Clean format for result(s)
/       (CleanFormatforPastedRunResults - Ctrl+Alt+Shift+1)
/   3.  Scroll beyond last result
/       (ScrollBeyondLastResult - Press down arrow in M3)
/   4.  Import latest result for each runner
/       (ImportLatestResultForEachRunner - Ctrl+Alt+Shift+9)
/   5.  Generate charts from Groups
/       (GenerateChartsFromGroupsSheet - Ctrl+Alt+Shift+0)
/ 	6.  Set legend colours for Groups
/       (SetLegendCellColoursOnSheet)
/---------------------------------------------------------------------------------------
 */

// @OnlyCurrentDoc
// @scope https://www.googleapis.com/auth/script.external_request

const resultTABLE = "Event"   // For any Runner Event results sheet
const eventHeaderCELL = "A2"; //  where the header row is below the Runner name title
const firstResultCELL = "A3"; //  and at least one Result has been cleanly entered
const firstResultROW = 3;     //  where new Results MUST be below this first result
const scrollCOLUMN = 12;      // Scroll down when selecting column M beyond the table
const pasteCOLUMN = 1;        // Paste Event result starting in 1st column (A)
const timeCOLUMN = 5;         // Event result time is in the 5th column (E)

// Junior parkrun thresholds?
const min2kmTIME = 6;         // Minimum time for 2km in minutes (after repair)
const max2kmTIME = 23;        // Maximum time for 2km (less than 24 "hours" = minutes)
// Table designed for 5k run times
const min5kmTIME = 13;        // Minimum time for 5km in minutes (after repair)
const max5kmTIME = 37;        // Maximum time for 5km (24 "hours"" after min5kmTIME)
// If ever for 10km runs, needs adapting to use display times (instead of date values)
const min10kmTIME = 27;       // Minimum time for 10km in minutes  (after repair)
const max10kmTIME = 51;       // Maximum time for 10km (24 "hours" after min10kmTIME)
const recentYRS = 3;          // filter comparison graphs based to most recent years only

/* ---------------------------------------------------------------------------
/
/   The following definitions and functions support the automatic protection
/   of each runner's result sheets to allow safe entry of results by others.
/   The heirarchy of functions are:
/     ReprotectEachRunnerResultsSheets
/         CreateRunnerResultsSheet (if new runner)
/         EnsureBlankResultsRange (range for new results)
/       ReallowRunnerResultsSheetWithException
/         UnprotectResultsSheet
/         ProtectResultsSheetWithException
/       ReprotectResultsRange
/         UnprotectResultsRangeOnSheet
/         GetResultsAdders
/         ProtectResultsRangeByEditor
/
/  ---------------------------------------------------------------------------
*/

const allRunnersSHEET =           // The Runners sheet drives the creation of results sheets
  SpreadsheetApp.getActiveSpreadsheet()   // ...for new runners by those permitted
    .getSheetByName("Runners");
const templateNameCELL = "J1";          // This cell identifes the seed template (e.g. Keren)
const templateNAME = allRunnersSHEET    // The seed template may be readily reconfigured
  .getRange(templateNameCELL)     // See the note on Runners!J1 cell
  .getValue();                    //  ...where only the first 3 (or 4) rows are relevant
const numBlankROWS = 10;          // Regular catch-up of multiple results by those permitted
const parkrunnerIdCOLUMN = 10;    // in column L on Runners sheet
const runnersStartROW = 3;      // start after title & header rows (2)

/**
 * Ensures a specified number of blank rows exist at the end of the active sheet.
 *  If necessary, inserts new rows to accommodate the specified number.
 *    @param {number} numRows - The number of blank rows to ensure (default: 10)
 *  @return {string} The A1 notation of the range of blank rows
 */
function EnsureBlankResultsRange(numRows=numBlankROWS) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var runnerName = sheet.getName();
  var resultRow = sheet.getLastRow();   // last non-blank row
  var maxRows = sheet.getMaxRows();     // actual final row number
  var numRowsToInsert = numRows-(maxRows-resultRow);
  if (numRowsToInsert > 0)
    sheet.insertRowsAfter(maxRows,numRowsToInsert);
  var endColumn = sheet.getLastColumn();
  var clearRange = sheet.getRange(resultRow+1,1,numRows,endColumn);
  clearRange.clearContent();
  var rangeNotation = clearRange.getA1Notation();
  var logMessage = "Range for allowing results to be added is "+rangeNotation;
  if (numRowsToInsert > 0)
    logMessage += " (with "+numRowsToInsert+" rows appended)";
  logMessage += " on "+runnerName + "'s results sheet";
  Logger.log(logMessage);
  return rangeNotation;  var clearRange = sheet.getRange(resultRow+1,1,numRows,endColumn).clearContent();;
  var rangeNotation = clearRange.getA1Notation();
}

/**
 * Unprotect entire sheet (including exceptions)
 */
function UnprotectResultsSheet(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  if (protections.length > 0) protections[0].remove();
}

/**
 * Protect entire sheet with exception (pending permit)
 */
function ProtectResultsSheetWithException(sheet,exceptionRange) {
  var sheetProtection = sheet.protect();
  // Only the worksheet owner is permitted to allow this range exception 
  var protectionRange = sheet.getRange(exceptionRange);
  // Exception range to be restricted thereafter
  sheetProtection.setUnprotectedRanges([protectionRange]);
}

/**
 * Reallows editing on a specific range of a protected sheet.
 *  Removes existing protection, adds an exception for the specified range,
 *  and reapplies protection with domain edit disabled.
 *    @param {string} rangeNotation - The A1 notation of the range to allow editing.
 */
function ReallowRunnerResultsSheetWithException(rangeNotation) {
  var sheet = SpreadsheetApp.getActiveSheet();
  UnprotectResultsSheet(sheet);
  ProtectResultsSheetWithException(sheet,rangeNotation);
  var runnerName = sheet.getName();
  Logger.log("Reprotected "+runnerName+"'s results except for the new results range, "+rangeNotation);
}

/**
 * Gets the list of users permitted to add results for a specific runner.
 *    @param {string} runnerName - The name of the runner.
 *  @return {string[]} An array of user names permitted to add results.
 */
function GetResultsAddersForRunner(runnerName) {
  // var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const addersCOLUMN = "D";     // Where the results adder(s) are in the Runners table
  var indexRow = allRunnersSHEET.getRange("A3:A").getValues().map(function(value) { 
    return value[0];
  }).indexOf(runnerName);
  if (indexRow == -1) {
    Logger.log("Individual runner, "+runnerName+" not found within the Runners sheet table");
    return [];
  }
  var addersCell = allRunnersSHEET.getRange(addersCOLUMN+(indexRow+3));
  var adders = addersCell.getValue().split(",");
  adders = adders.map(adder => adder.trim());
  Logger.log("Results on " + runnerName + "'s result sheet may be added by "
    +adders.join(", "));
  return adders;
}

/**
 * Remove any range protections on a runner's sheet (typically one only)
 */
function UnprotectResultsRangeOnSheet(sheet) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function(protection) {
    protection.remove();
  });
}

/**
 * Allow the new range protection on a runner's sheet
 */
function ProtectResultsRangeByEditor(editor,newProtection) {
  try {
    newProtection.addEditor(editor);
  } catch (err) {
    Logger.log("Error adding editor: "+editor+". Error: "+err+" because not a google email address");
  }
}

/**
 * Reprotects a specific range of a sheet, allowing only specified users to edit.
 *    @param {string} rangeNotation - The A1 notation of the range to protect.
 *  @return {string[]} An array of user names permitted to edit the range.
 */
function ReprotectResultsRange(rangeNotation) {
  var sheet = SpreadsheetApp.getActiveSheet();
  UnprotectResultsRangeOnSheet(sheet);
  var protectionRange = sheet.getRange(rangeNotation);
  var newProtection = protectionRange.protect();
  var runnerName = sheet.getName();
  var adders = GetResultsAddersForRunner(runnerName);
  if (adders && adders.length > 0) {
    adders.forEach(function(adder) {
      ProtectResultsRangeByEditor(adder,newProtection);
    });
    Logger.log("Reprotected new range, "+rangeNotation+" on "+runnerName+
      "'s results sheet, except by "+(adders.join(", ")));
  } else {
    SpreadsheetApp.getUi().alert('Warning',"Without protection, any spreadsheet Editor can add results to "+
      runnerName+"'s results sheet");
    Logger.log("Unprotected new range, "+rangeNotation+" on "+runnerName+"'s results sheet");
  }
  return adders;
}

/**
 * Creates a new results sheet for a runner if one doesn't exist.
 *    @param {Array of strings} runnerFullName - The full name of the runner
 *    @param {number} parkrunnerId - The Park Runner ID
 * @return {Sheet} The new results sheet
 */
function CreateRunnerResultsSheet(
  runnerFullName = ["Joe","Hodgson"],  // Allow for test case!
  parkrunnerId = 11610909)
{
  // Create runner's results sheet if it doesn't exist
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var templateResults = spreadsheet.getSheetByName(templateNAME);
  var runnerName = runnerFullName[0];
  var newResultsSheet = templateResults.copyTo(spreadsheet).setName(runnerName);
  // ensure the content of the sheet is unique
  var fullName = runnerFullName.join(" ");
  const titleNameCELL = "A1";
  const parkrunnerIdCELL = "L1";
  const endTemplateROW = 4;
  newResultsSheet.getRange(titleNameCELL).setValue(fullName);
  newResultsSheet.getRange(parkrunnerIdCELL).setValue(parkrunnerId);
  // trim table leaving a phantom result row as a model for adding more
  var numRows = newResultsSheet.getLastRow()-endTemplateROW;
  if (numRows > 0) newResultsSheet.deleteRows(endTemplateROW+1,numRows);
  return newResultsSheet;
}

/** 
 * Reprotects each runner's results sheet, to allow permitted users to add results.
 *  For each named runner...
 *    0. Get that runner's results sheet (or create if none)
 *    1. Ensure sufficient blank rows to add new result(s)
 *    2. Protect that results sheet except for the blank row range
 *    3. Protect that blank range for specific (editor) runners only to add results
 */
function ReprotectEachRunnerResultsSheets() {
  const runnersRANGE = "A"+runnersStartROW+":B";
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var runnerFullNames = allRunnersSHEET.getRange(runnersRANGE)
    .getValues().filter(function(value) {
    return value[0] != "";
  }).map(function(value) {
    return [value[0],value[1]]; 
  });
  runnerFullNames.forEach(function(
    runnerFullName,index)
  {
    var runnerName = runnerFullName[0];
    var resultsSheet = spreadsheet.getSheetByName(runnerName);
    if (!resultsSheet) {
      var thisRow = index+runners1stROW;
      var parkrunnerId = allRunnersSHEET.getRange(
        thisRow,parkrunnerIdCOLUMN).getValue();
      resultsSheet = CreateRunnerResultsSheet(runnerFullName,parkrunnerId);
      if (!resultsSheet) return;
    } 
    resultsSheet.activate();
    var rangeNotation = EnsureBlankResultsRange(numBlankROWS);
    ReallowRunnerResultsSheetWithException(rangeNotation);
    var adders = ReprotectResultsRange(rangeNotation);
  });
  Logger.log("Protection complete on each runner's results sheet");
}

/* ----------------------------------------------------------------------------
/
/   The following definitions and functions support the efficient manual entry
/   and "cleaning" of new results into the (permitted) runner's results sheet..
/
/  ----------------------------------------------------------------------------
*/

function ScrollBeyondLastResult() {
  const functionNAME = "ScrollBeyondLastResult";
  // Select the first empty row in preparation for pasting Event results
  // Conveniently used to skip to the bottom of the table below the precedent result
  var sheet = SpreadsheetApp.getActiveSheet();
  Logger.log('Running '+functionNAME+' on sheet: '+sheet.getName());
  var lastRow = sheet.getLastRow();
  if (lastRow <= firstResultROW) return; // table has no results yet
  // Find an empty row from the bottom of the table since more efficient
  var i = lastRow;
  for (; i>firstResultROW; i--) 
    if (sheet.getRange(i,pasteCOLUMN).getValue()) break;
  // When last row has a result, prepare for pasting new result(s) below
  if (i == lastRow)       
    sheet.appendRow([""]);
  // Always select the row below the last result
  sheet.getRange(i+1,pasteCOLUMN).activate();
}

/**
 * Sets up skipping to the bottom (on selecting L3 cell)
 *  Instructions to set up (and execute?) the trigger:
 *    1. Go to the Apps Script (Macros) editor.
 *    2. Click on the clock icon (Triggers) in the left sidebar.
 *    3. Click on "Create trigger".
 *    4. Set up the trigger with the following settings:
 *        - Choose function: onSelectionChange
 *        - Select event type: On select
 *        - Save
 */
function onSelectionChange(e) { 
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  // Ignore sheet quickly if table does not contain Event results
  if (sheet.getRange(eventHeaderCELL).getValue() !== resultTABLE) return;
  // Skip also if first result has not been entered (and cleaned manually)
  if (sheet.getRange(firstResultCELL).getValue() === "") return;
  // Only effective when selecting a cell just to the right of the table
  // assuming a down-arrow image is on the right of that Event result header
  if (range.getColumn() === scrollCOLUMN)
    ScrollBeyondLastResult();
}

function convertHoursToMinutes(timeStr) {
  // Results are 60 times slower, hh:mm and need to be mm:ss
  var showMinutes = timeStr.toString().split(':');
  Logger.log('Minutes shown is '+showMinutes+' in the cell for string, '+timeStr); 
  return parseInt(showMinutes[0])/60;
}

/**
 * Clears the borders from the specified range.
 *    @param {Range} range - The range to apply the format to
 */
function ClearBordersOnRange(range) {
  var firstRow = range.getRow();
  var lastRow = range.getLastRow();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  range.setBorder(false,false,false,false,false,false);
  Logger.log('1. Cleared borders on new row(s), '+firstRow+' to '+lastRow+
    ' (for '+numRows+' rows and '+numCols+' columns)');
}

/**
 * Applies the format from the row above to the specified range.
 *    @param {Range} range - The range to apply the format to
 */
function ApplyFormatFromRowAboveOnRange(range) {
  var firstRow = range.getRow();
  var firstCol = range.getColumn();
  var prevRow = firstRow-1;
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var aboveRange = range.offset(-1,0,numRows,numCols);
  var lastRow = firstRow+numRows-1;
  var lastCol = firstCol+numCols-1;
  var sheet = range.getSheet();
  aboveRange.copyFormatToRange(sheet,firstCol,lastCol,firstRow,lastRow);
  Logger.log('2. Formatted the new row(s), '+firstRow+' to '+lastRow+
      ' (for '+numRows+' rows and '+numCols+' columns)'+
      ' based on format in above row, '+prevRow);
}

/**
 * Applies formulas (including any values) and data validation from the row above
 *  to the specified range, beyond the original range.
 *    @param {Range} range - The range to extend formulas and data validation from
 */
function ApplyFormulaFromRowAboveBeyondRange(range) {
  const prevRowOFFSET = -1;
  var firstRow = range.getRow();
  var firstCol = range.getColumn();  // typically column 1 (A) up to column G (7) pasted
  var formCol = range.getLastColumn()+1; // next column, typically column 8 (H)
  var numRows = range.getNumRows();
  var endColumn = range.getSheet().getLastColumn(); // sheet end Column (with data), typically 13 (M)
    // typically offset by 7 to column H upto end column M (13) means 6 columns
  var offsetCol = formCol-firstCol;
  var numCols = endColumn-offsetCol; 
  var aboveRange = range.offset(prevRowOFFSET,
    offsetCol,1,numCols);  // modelled on one row
  var beyondRange = range.offset(0,
    offsetCol,numRows,numCols); // for new row(s)
  var formulae = aboveRange.getFormulas();
  // var lastRow = firstRow+numRows-1;
  aboveRange.copyTo(beyondRange,SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
  aboveRange.copyTo(beyondRange,SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION);
  Logger.log('3. Inherited formulae (incl, check boxes & drop-downs) from column, '+
      String.fromCharCode(64+formCol)+' (for '+numCols+' columns)'+
      ' from row, '+firstRow+' (for '+numRows+' rows)');
}

/**
 * Translates an elapsed duration time in hours as if minutes,
 *  e.g. 27:55:00.000 means 00:27:55. assuming time less than an hour
 *    @param {time}   date/time object value in hh:mm format
 *  @returns {duration} as a string, 00:mm:00 for correct re-entry
 */
function TranslateDurationFromHoursToMinutes(time) {
  var newMinutes = time.getHours();   // previously shifted left incorrectly 
  var newSeconds = time.getMinutes();
  // ASSUMPTION: sub-14 minutes close to world record!
  if (newMinutes < min5kmTIME)
    newMinutes += 24;  // Assume "date" is effectively 24 hours later!
  eHours = "00";
  eMinutes = newMinutes.toString().padStart(2,'0');
  eSeconds = newSeconds.toString().padStart(2,'0')+'.000';
  var durationStr = eHours+":"+eMinutes+":"+eSeconds;
  Logger.log('Duration translated from hh:mm to mm:ss duration as '+durationStr); 
  return durationStr;
}

/**
 * Repairs duration times in a specified column of a range,
 *  translating each value from hh:mm to 00:mm:ss (i.e. 60 times faster!)
 *    @param {Range} - The range containing the duration times to repair
 *    @param {number} [timeCol=timeCOLUMN] - Column containing the duration times
 */
function RepairDurationTimesInRangeColumn(range,
  timeCol = timeCOLUMN)
{
  var firstRow = range.getRow();
  var numRows = range.getNumRows();
  var lastRow = firstRow+numRows-1;
  // time pasted as hh:mm (instead of mm:ss
  var timeRange = range.offset(0,timeCol-1,numRows,1);  // offset from column A
  var timeValues = timeRange.getValues();
  // Prepend each results' duration with 00: (for hh:) since 20 hours means 20 minutes
  var timeRepairs = timeValues.map(function(row) {
    return [TranslateDurationFromHoursToMinutes(row[0])];
  });
  timeRange.setValues(timeRepairs);
  Logger.log('4. Repaired time(s) in column, '+
    String.fromCharCode(64+timeCol)+' for new row(s), '+
    firstRow+' to '+lastRow+' (for '+numRows+' rows),'+
    ' to be mm:ss duration, instead of hh:mm (60 times longer)');
}

/**
 * Clean up one or more pasted run results in 4 steps
 *    @param {Range} -optionally pre-selected range (otherwise null)
 *  Preconditions:
 *    - Result(s) from parkrun(s) pasted below existing clean result(s)
 *    - Earlier result is below the runner title & header rows
 *  Steps to follow before running this macro:
 *    A. Open the URL on the right of the first row (in M1)
 *        or from your Surname on the Runners sheet
 *    B. Scroll to All Results, select the entire row from your latest result, and Copy
 *        or, for multiple results, Sort by Date, select missing rows at the bottom, and Copy
 *    C. Skip to 1st empty row on your named sheet, and Paste result(s) into 1st column
 *        or press the green arrow button (in K3), and Paste into 1st empty cell
 *    D. Execute macro via Extension >> Macros >> Clean Format for Pasted Run Result(s)
 *        or press defined hotkey combination, Ctrl+Alt+Shift+1
 */
function CleanFormatforPastedRunResults(
  pastedResults)    // optional range, pre-selected
{
  const functionNAME = "CleanFormatforPastedRunResults";
  // Used to clean up one or more results in 4 steps (1..4), AFTER A, B & C beforehand 
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  Logger.log('Running '+functionNAME+' on sheet: '+sheet.getName());
  // Preconditions: ensure result(s) from parkrun(s) pasted below existing clean result(s)
    if (!pastedResults)   // if range specified, assume also already active
      pastedResults = sheet.getActiveRange();   // ...such that it should match
    var firstRow = pastedResults.getRow();
    if (firstRow <= firstResultROW) {
      SpreadsheetApp.getUi().alert('Error','Ensure that there is a precedent clean result immediately above the newly pasted result(s) to be cleaned (assuming that earlier result is below the runner title & header rows)',SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var pastedValue = sheet.getRange(firstRow,pasteCOLUMN).getValue();  
    if (!pastedValue) {
      SpreadsheetApp.getUi().alert('Error','Ensure that the result(s) have been copied from the All Results table'+
        '(with the Event Name and PB?) and that the pasted result(s) to be cleaned remain selected',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var pastedColumn = pastedResults.getColumn();
    if (pastedColumn !== pasteCOLUMN) {
      SpreadsheetApp.getUi().alert('Error','Ensure that the result(s) have been pasted into the first column (A) and that they remain selected',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;   
    }
    var firstTimeString = sheet.getRange(firstRow,timeCOLUMN).getDisplayValue();
    var firstElapsedTime = convertHoursToMinutes(firstTimeString);
    if (firstElapsedTime < min5kmTIME) {
      SpreadsheetApp.getUi().alert('Error', 'Elapsed time (after repair) expected to exceed world record for 5km',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  // Steps 1-4 handled by sub-functions:
  ClearBordersOnRange(pastedResults);
  ApplyFormatFromRowAboveOnRange(pastedResults);
  ApplyFormulaFromRowAboveBeyondRange(pastedResults);
  RepairDurationTimesInRangeColumn(pastedResults,timeCOLUMN);
}

/* --------------------------------------------------------------------------
/
/   The following definitions and functions automate the addition of a new
/   result by securely connecting to the parkrun site for importing the
/   latest result, mimicking a manual extract to extract copy and paste.
/   The heirarchy of functions are:
/
/       ImportLatestResultForEachRunner
/         OpenChromeBrowser ->
//        Loop for each runner...
/           >CopyLatestResultForRunner ->
/             GetParkrunnerResultsPage
/               AccessPage
//          When new result...
/             >PasteLatestResultForRunner ->
//              Loop for each cell...
/                 CleanValue (extracts links)
/               FormatDate
/               ApplyLinks
/             >AppendPositionsForRunner (with results link) ->
/               ExtendRange
/               >AssessPositions
/               IncludePositions
/         >CloseChromeBrowser
/
/  -------------------------------------------------------------------------
*/

/**
 * For the purpose of this trial, the following session & connection functions rely on
 * The asynchronous functions include:
 *    OpenChromeBrowser
 *    AccessPage
 *    CloseChromeBrowser
 */

const browserURL = 'https://browser-automation-service-224251628103.europe-west1.run.app';    // Google Cloud service in operation
const sampleURL = 'https://www.example.com';  // default test

async function OpenChromeBrowser() {
  const initBrowserURL = browserURL+'/initBrowser';
  try {
    var response = await UrlFetchApp.fetch(initBrowserURL);
    Logger.log(response.getContentText());
  } catch (err) {
    Logger.log(err);
  }
}

async function AccessPage(
  thisUrl = sampleURL
) {
  let getBrowserURL = browserURL+'/getUrl?url='+thisUrl;
  try {
    var response = await UrlFetchApp.fetch(getBrowserURL,
      // Runners with 500+ results take longer 
      //    AND parallelism may be curtailed
      //      WHEN re-using the same page is detected,
      //      AND/OR WHEN a single CPU is specified for the run
      // Therefore, allow worst case timing for all, as if serial
      {muteHttpExceptions:true,timeout:30000});
    return response.getContentText();
  } catch (err) {
    Logger.log(err,'within',thisUrl);
    return null;
  }
}

async function CloseChromeBrowser() {
  const stopBrowserURL = browserURL+'/stopBrowser';
  try {
    var response = await UrlFetchApp.fetch(stopBrowserURL);
  	Logger.log(response.getContentText());
  } catch (err) {
    Logger.log(err);
  }
}

const parkrunURL = 'https://www.parkrun.org.uk';
const parkrunnerURL = parkrunURL+'/parkrunner/';

function GetParkrunnerResultsPage(
  parkrunnerId = '777764')    // default useful for testing
{
  thisParkrunnerURL = parkrunnerURL+parkrunnerId +'/all/';
  return AccessPage(thisParkrunnerURL)
  .then (htmlContent => {  // ensures htmlContent is not a Promise
    return htmlContent;
  })
  .catch (error => {
    Logger.log('ERROR: Unable to access page, '+thisParkrunnerURL+'...\n'+error);
    return null;
  });
}

/**
 * Copies the single latest result for a given parkrunner ID from parkrun.org.uk
 *    @param {string} parkrunnerId - The ID of the parkrunner (default: "777764")
 *  @returns {array|null} - represents the latest result (or null if none)
 */
async function CopyLatestResultForRunner(
  parkrunnerId = "777764")
{
  const allResultsTITLE = "All  Results";    // 3rd table with All Results
  var headerStart = "<thead><tr><th>Event</th><th>Run Date</th><th>Run Number</th><th>Pos</th>";
  const allHEADER = headerStart+"<th>Time</th><th>Age<br/>Grade</th><th>PB?</th></tr></thead>";
  // WARNING: Row does not (yet) show Gender position,
  //          nor positions for Age-Category or Age-Grade (%age)
  return GetParkrunnerResultsPage(parkrunnerId)
  .then (allResults => {   // ensures allResults is not a Promise
    if (allResults) {
      allResults = allResults
        .substring(allResults
          .indexOf(allResultsTITLE))  // trim before "All  Results"
        .split('#comments')[0];      // trim after #comments
      if (allResults) {
        // Logger.log('All Results table: '+allResults); 
        var headerRow = allResults.match(/<thead>.*?<\/thead>/s);
        if (headerRow && headerRow.length>0) {  
          // Logger.log('Header row: '+headerRow[0]);  // TODO: columns may change - WARN if so?
          var bodyContent = allResults.match(/<tbody>.*?<\/tbody>/s)[0];
          var bodyRows = bodyContent.match(/<tr[^>]*>.*?<\/tr>/gs);
          if (bodyRows && bodyRows.length>0) {
            const latestROW = 0;  // only consider 1st row as single latest result
            var resultRow = bodyRows[latestROW];
            var cells = resultRow.match(/<td>.*?<\/td>/gs);
            if (cells) {
              var values = cells.map(function(cell) {
                return cell.replace(/<td>|<\/td>/g,"").trim();
              });
              Logger.log('Cells parsed as '+cells+' for runner, '+parkrunnerId+', and ready for pasting');
              return values;
            } else {
              Logger.log('WARNING: No cells found in row, '+resultROW+' for runner, '+parkrunnerId);
              return null;
            }
          } else {
            Logger.log('WARNING: Unable to find row, '+resultROW+' in all results for runner, '+parkrunnerId);
            return null;
          }
        } else {
          Logger.log('WARNING: Missing all results table for runner Id, '+parkrunnerId);
          return null;
        }
      } else {
         Logger.log('WARNING: No results yet for runner Id, '+parkrunnerId);
        return null;
      }
    } else {
      Logger.log('ERROR: Missing results page for runner Id, '+parkrunnerId);
      return null;
    }
  })
  .catch (error => {
    Logger.log('ERROR: Failed to get results for runner Id, '+parkrunnerId+'...\n'+error);
    return null;
  });
}

/**
 * Pastes the latest result for a given runner into their sheet.
 *    @param {array} latestResult - representing the latest (single) result
 *    @param {string} [runnerName="Alan"] - name of the runner's results sheet
 *  @returns {Range} - where result was pasted (or null if result already pasted)
 */
 async function PasteLatestResultForRunner(
  latestResult,         // assumed a single row to be selected
  runnerName = "Alan")
{
  let hyperLinks = [];

  /**
   *  Clean (each) cell, perhaps double clean if necessary)
   *  ...while capturing the href hyperlinks from the HTML (to apply later to the cell)
   *    @param {string} cell - Cell value to clean
   *  @returns {string|number} Cleaned cell value
   *  @sideeffect Pushes hyperlinks to global hyperLinks array
   */
  function CleanValue(cell) {
    var value = cell;
    switch (true) {
      case cell.startsWith('<span '):       // for Date, ALSO has an outer <a href
        value = cell.replace(/<span[^>]*>(.*?)<\/span>/,'$1');
      case cell.startsWith('<a href='):   // for Event (location), Date, or Run #
        var href = value.match(/href=['"]([^'"]+)['"]/)[1];
        hyperLinks.push(href.replace(/<[^>]*>/g,''));
        value = value.replace(/<[^>]*>/g,'');
        return value;
      case /^\d{2}:\d{2}$/.test(cell):      // for Elapsed time (under an hour)
        return '00:' + value;
      case cell.includes(':'):              // for Elapsed time (over an hour)
        return value;
      case cell.includes('%'):              // for Age-grade (%age)
        return value;
      case /^\d+$/.test(cell):              // Posn (potentially other numbers)
        return parseInt(value,10);
      case cell.includes('&nbsp'):
        value = cell.replace(/&nbsp;/g,''); // for PB?, ALSO needs trim
      default:
        return value.trim();
    }
  }

  /**
   *  Formats a date string into a specified format.
   *    @param {string} dateSource - Date string to format (DD/MM/YYYY expected)
   *    @param {string} dateFormat - Output date format (e.g. "d-MMM-yyyy")
   *  @returns {string} Formatted date string
   */
  function FormatDate(dateSource,dateFormat) {
    return Utilities.formatDate(  // ensure ISO dates beforehand
      new Date(dateSource.replace(/(\d+)\/(\d+)\/(\d+)/,'$3-$2-$1')),
      Session.getScriptTimeZone(),
      dateFormat  // more readable & universal, PLUS consistent with current formatting on sheet
    );
  }

  /**
   *  Applies hyperlinks (previously extracted from latest result) into results sheet.
   *    @param {string[]} latestResult - Latest result (row cell values) for a runner
   *    @param {Sheet} resultsSheet - Sheet to append the latest result row
   *    @param {string} runnerName - Runner's name
   *  @returns {string} resultsLink - link to results sheet to determine detailed positions
  */
  function ApplyLinks(latestResult,resultsSheet,runnerName) {
    var insertRow = resultsSheet.getLastRow()+1;
    try {
      latestResult = latestResult.map((cell,i) => {
        // Perhaps assume only columns A..C had hyperlinks (captured during CleanValue)
        // if ([0,1,2].includes(i)) {
        if (hyperLinks.length) {
          return SpreadsheetApp.newRichTextValue()
            .setText(cell)
            .setLinkUrl(hyperLinks.shift())   // reduces length of array
            .build();
        } else {
          return SpreadsheetApp.newRichTextValue()
            .setText(cell)
            .build();
        }
      });
    } catch (error) {
      Logger.warn('WARNING: Unable to add hyperlinks correctly to latest result for runner, '+runnerName+'...\n'+error.message);
    }
    var pastedResult = resultsSheet.getRange(insertRow,1,1,latestResult.length)
      .setRichTextValues([latestResult]); // append single result row only (at column A)
    // SpreadsheetApp.flush();   // ensure values are visible for retrieving results = no effect
    return pastedResult;
  }

  // Logger.log('Consider pasting latest result into runner sheet for '+runnerName+'...');
  const locationINDEX = 0;  // column A
  const dateINDEX = 1;      // column B
  const previousROWS = 15;  // previous results may include non-parkrunner events
  const dateFORMAT = 'd-MMM-yy';
  latestResult = latestResult.map(CleanValue);   // includes 00: (for hh:) prefix on elapsed time
  var thisDate = latestResult[dateINDEX] = FormatDate(latestResult[dateINDEX],dateFORMAT);
  var thisLocation = latestResult[locationINDEX];
  var resultsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(runnerName);
  // Consider matching only with recent previous results (after ignoring title & header rows)
  // This takes account of some runners who run elsewhere (and captured manually) since their last recorded parkrun
  var previousResults = resultsSheet.getDataRange().getDisplayValues().slice(2).slice(-previousROWS);
  if (previousResults.length < 1) {
    Logger.warn('WARNING: The template for new runner, '+runnerName+' is expected to include a (phantom?) formatted result');
    Logger.error('ERROR: No previous result for runner, '+runnerName+' (with map formulae?) to be able to match new result');
    return null;
  }
  var matchingResults = previousResults.map(function(row) {
    return row[locationINDEX]+'&'+row[dateINDEX];
  });
  if (matchingResults.includes(thisLocation+'&'+thisDate)) {
    Logger.log('Previously pasted result at '+thisLocation+' on '+thisDate+' for runner, '+runnerName);
    return null;    // null range if not a new result
  } else {
    var resultsLink = hyperLinks[dateINDEX];  // Retain the link to the event results before consumed by the sheet
    var pastedResult = ApplyLinks(latestResult,resultsSheet,runnerName);
    Logger.log('Newly pasted result at '+thisLocation+' on '+thisDate+' for runner, '+runnerName);
    // WARNING: Do NOT rely on any hyperlink being stored on the sheet (yet)
    return {
      resultRange: pastedResult,
      resultsLink: resultsLink  // To determine detailed positions (not reliant on retrieving from the sheet)
    };
  }       
}

/**
 * Determines extra positions for a runner from parkrun results page
 *  @param {string} thisUrl 
 *  @param {string} matchRunner
 *  @param {string} ageCat 
 * @return {Promise<Array>} [acPosn,agPosn,gcPosn]
 */
async function AssessPositions(thisUrl,matchRunner,ageCat,genderCat) {
  let filterBrowserURL = browserURL+'/filterUrl'
    +'?url='+thisUrl
    +'&rn='+encodeURIComponent(matchRunner)
    +'&ac='+ageCat
    +'&gc='+genderCat;
    // +'&ag='Age-Grade';   // assume internal default
  try {
    var positions = await UrlFetchApp.fetch(filterBrowserURL,
      // Runners with 500+ results take longer 
      //    AND parallelism may be curtailed
      //      WHEN re-using the same page is detected,
      //      AND/OR WHEN a single CPU is specified for the run
      // Therefore, allow worst case timing for all, as if serial
      {muteHttpExceptions:true,timeout:40000});
    Logger.log(positions.getContentText());
    var posns = JSON.parse(positions);
    return [posns.acPosition,posns.agPosition,posns.gcPosition];
  } catch (err) {
    Logger.log(err,'within',thisUrl);
    return null;
  }
}

/**
 * Includes extra positions into latest runner's result row on sheet
 *  @param {Range} resultRange 
 *  @param {string} acPosn 
 *  @param {string} agPosn
 *  @param {string} gcPosn (replacing redundant run #)
 */                                                 
function IncludePositions(resultRange,acPosn,agPosn,gcPosn) {
  const GenderPosnCOL = 9;      // column I (new column extension)
  const AgeCatPosnCOL = 10;     // column J (previously done manually)
  const AgeGradePosnCOL = 11;   // column K (new column extension)
  resultRange.getCell(1,AgeCatPosnCOL).setValue(acPosn);    // col J for Age-Cat position
  resultRange.getCell(1,AgeGradePosnCOL).setValue(agPosn);  // col K for Age-Grade position
  // Pending outcome of proposal to Parkrun site: "Replace the Run # with Gender Position?
  resultRange.getCell(1,GenderPosnCOL).setValue(gcPosn);    // col I for Gender position
  // return resultRange;   // Values change, although previously extended range itself is unchanged
}

/**
 * Extends the runner's result range, typically one result in a single row
 *  @param {Range} resultRange 
 *  @param {numeric} numRows 
 *  @param {string} agPosn
 */
function ExtendRange(resultRange,numRows,numCols) {
  const PBtickBoxCOL = 8; // column H is a read-only tick-box, derived from column G (if PB present)
  const AgeCatCOL = 12;   // column L is the age-category, derived from event date  (without conflict)
  // Two derived values are generated from a MAP function set for the first result in the table 
  resultRange = resultRange.offset(0,0,resultRange.getNumRows()+numRows,resultRange.getNumColumns()+numCols);
  resultRange.getCell(resultRange.getNumRows(),PBtickBoxCOL)
    .setValue(null);    // required to prevent FALSE if not "PB" in col G
  // resultRange.getCell(resultRange.getNumRows(),AgeCatCOL)
  //  .setValue(null);  // no conflict in derivation from event date in col B and runner's DoB
  Logger.log('Extended range of results by '+numRows+' rows and '+numCols+' columns');
  // resultRange implicitly includes the derived Age-Category in col L (assumes MAP function in 1st result)
  var ageCatCell = resultRange.getCell(1,AgeCatCOL);
  var ageCatCellRef = ageCatCell.getA1Notation();
  var ageCat = ageCatCell.getValue();
  Logger.log('Age-Category is '+ageCat+' in new runner row at cell, '+ageCatCellRef);
  return resultRange;
}

/**
 * Appends detailed positions to an individual runner's sheet beyond the current range
 */
async function AppendPositionsForRunner(
  runnerName,surname,genderCat,
  resultRange,resultsLink)
{
  const eventCOL = 1;
  const dateCOL = 2;
  const AgeCatCOL = 12;      // column L
  const PBtoAgeCatNumCOLS = 5;   // cols H..L for I..K positions between derived cols, H & L
  var runnerFullName = runnerName+' '+surname;
  var eventDateCell = resultRange.getCell(1,dateCOL);
  var eventDate = eventDateCell.getDisplayValue();
  var eventLocation = resultRange.getCell(1,eventCOL).getValue();
  Logger.log('New date is '+eventDate+' at '+eventLocation+' for runner, '+runnerName);
  var extResultRange = ExtendRange(resultRange,0,PBtoAgeCatNumCOLS);  // extends to include H..K
  var ageCat = extResultRange.getCell(1,AgeCatCOL).getValue();  // derived from Date - DoB 
  Logger.log('Age-Category is '+ageCat+' for runner, '+runnerName);
  // const genderCat is fixed per runner from the full Runners sheet
  // WARNING: Expected resultsLink not (yet) accessible from the extended range!!
  // var resultsLink = eventDateCell.getRichTextValue().getLinkUrl();  // behind Date (in col B)
  Logger.log('Link to '+eventLocation+' latest results: '+resultsLink);
  return AssessPositions(resultsLink,runnerFullName,ageCat,genderCat).then(extraPosns => {
    if (extraPosns.length === 3) {
      IncludePositions(extResultRange,...extraPosns);  // into cells, I..K on result row
      Logger.log('Positions for Gender & Age Category, plus Age-Grade appended to latest result in runner sheet, '+runnerName);
    } else {
      Logger.error('ERROR: Unable to append positions for latest result in runner sheet, '+runnerName);
    }
  });
}
/**
 * Imports the latest results for each of our runners from parkrun.org.uk.
 */
function ImportLatestResultForEachRunner() {
  OpenChromeBrowser().then(() => {   // Browser must have launched beforehand...
    var runners = allRunnersSHEET.getRange(
      "A"+runnersStartROW+":C"
    ).getValues().filter(String);
    // Process each runner in parallel BEFORE closing after ALL runners done!
    Promise.all(runners.map(function(runner,index) {
    // runners.forEach(function(runner,index) {   // risky since all in parallel
      var runnerName = runner[0];     // col A
      var runnerSurname = runner[1];  // col B
      var runnerGender = runner[2];   // col C
      var parkrunnerId = allRunnersSHEET.getRange(
        runnersStartROW+index,parkrunnerIdCOLUMN).getValue();
      Logger.log('Parkrunner ID: '+parkrunnerId);
      if (isNaN(parkrunnerId)) {
        Logger.log('WARNING: Skipping invalid parkrunner Id, '+parkrunnerId);
        return; // since continue is n/a within forEach
      }
      // Ensure ALL functions complete by returning the Promise chain
      return CopyLatestResultForRunner(parkrunnerId).then(latestResult => {
        if (latestResult) {
          return PasteLatestResultForRunner(latestResult,runnerName).then(result => {
            if (result && result.resultRange && result.resultsLink) {
              return AppendPositionsForRunner(runnerName,runnerSurname,runnerGender,
                result.resultRange,result.resultsLink);
              Logger.log('Latest result appended to runner sheet, '+runnerName);
            } else {
              Logger.log('No later result needs appended for runner, '+runnerName);
            }
          });
        } else {
          Logger.log('WARNING: No results found for runner, '+runnerName);
        }
      });
    })).then(() => {
      return CloseChromeBrowser().then(() => {}); // Ensure all logs complete for each runner
    });
  });
}

/* ---------------------------------------------------------------------------
/
/   The following definitions and functions are used to support the automatic
/   creation of comparison charts /**
 * Imports the latest results for each of our runners from parkrun.org.uk.
 based on pe-defined groups of runner,
/
/ ---------------------------------------------------------------------------
*/

const chartYAxisTITLE = 'Age Grade';
const ageGradesCOL = 5;       // for column F on each runner's results sheet
const ageGradeFORMAT = '0.0%';  // for %age on graphs
const dateFORMAT = 'd-MMM-yy';

/**
 * Returns an array of selective runners performances versus event dates,
 *  ensuring event dates are unique (even if event location differs),
 *  based on most recent years (unless all performances override).
 *    @param {Array<Date>} allDates     - Array of non-unique event dates
 *    @param {Array<string>} runners    - Array of selective runners' names
 *    @param {Object} runnersPerfs      - Object of performances for each runner on dates (subset of all dates)
 *    @param {number} [mostRecentYears=recentYRS] - Number of recent years to include (0 to get all)
 *  @returns {Array<Array>} as an Array of arrays containing dated performances (with nulls for absences)
 *           (potentially more efficient as a 2D array)
 */
function CollateRunnersDatedPerformances(allDates,runners,runnersPerfs,
  mostRecentYears = recentYRS)  // assume all performances if zero years cut-off
{
  // Ensure empty performances for runners who missed out
  //    on any run dates (based on recent years only)
  var filteredDates = mostRecentYears == 0   // ignore filter option
    ? allDates
    : allDates.filter(date => date >= new Date(
      new Date().getFullYear()-(mostRecentYears|0),
      new Date().getMonth()-((mostRecentYears%1)*12|0),
      new Date().getDate()));
  var uniqueDates = [...new Set(filteredDates.map(date => date.getTime()))]
    .map(timestamp => new Date(timestamp)).sort((a, b) => a - b);
  var runnersDatedPerfs = [];
  for (var i=0; i<uniqueDates.length; i++) {
    var row = [uniqueDates[i]];
    for (var j=0; j<runners.length; j++) {
      if (runnersPerfs[runners[j]] &&
          runnersPerfs[runners[j]][uniqueDates[i]] !== undefined) {
        row.push(runnersPerfs[runners[j]][uniqueDates[i]]);
      } else {
        row.push(null);
      }
    }
    runnersDatedPerfs.push(row);
  }
  return runnersDatedPerfs;
}

/**
 * Retrieves and collates dated performances for the specified runners from their results sheets
 *    @param {Array<string>}                      - runners - Array of runner names
 *    @param {number} [mostRecentYears=recentYRS] - Number of recent years to include (0 to get all)
 *    @param {number} [perfColumn=ageGradesCOL]   - Column index of performance values (Age Grade, Time, etc.)
 *  @returns {Array<Array>} Array of arrays containing dated performances (with nulls for absences)
 */
function GetRunnersDatedPerformances(runners,
  mostRecentYears = recentYRS,   // assume all performances if zero years cut-off
  // performance is normally Age Grade values < 1, presented as %ages,
  //   but potentially Time or Age Grade Posn or # of 1sts may be used
  perfColumn = ageGradesCOL)
{
  const datesCOL = 1;     // for column B on each runner's sheet
  const headerROW = 2;    // number of header rows assumed above runner's results
  var allDates = [];
  var runnersPerfs = {};
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  runners.forEach(function(runnerName,index) {
    try {
      var resultsSheet = spreadsheet.getSheetByName(runnerName);
      if (!resultsSheet) {
        SpreadsheetApp.getUi().alert('Error',
          'No results sheet for indexed ('+index+') runner named, '+runnerName,
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      var resultsRange = resultsSheet.getDataRange();
      var values = resultsRange.getValues();  // dates in UTC & %ages < 1
      var dates = values.map(function(row) {
        return row[datesCOL]; // raw Date values
      }).slice(headerROW);
      var perfs = values.map(function(row) {
        return row[perfColumn];
      }).slice(headerROW);
      runnersPerfs[runnerName] = {};
      for (var i=0; i<dates.length; i++) {
        runnersPerfs[runnerName][dates[i]] = perfs[i];
      }
      allDates = allDates.concat(dates);   // dates are not unique
    } catch (e) {
      Logger.log("Error protecting runner sheet, " +runnerFullName[0]+": "+e);
    }
  });
  runnersDatedPerfs = CollateRunnersDatedPerformances(
    allDates,runners,runnersPerfs,mostRecentYears);
  Logger.log('GetRunnersDatedPerformances: '+runners);
  return runnersDatedPerfs;
}

function TransposeArray(array) {
  return array[0].map((_, colIndex) => array.map(row => row[colIndex]));
}

/**
 * Copies runners' performances to an existing sheet, formatting dates and values as specified.
 *    @param {Sheet} perfsSheet                 - An existing sheet where runners' performances are to be placed
 *    @param {Array<string>} runners            - Array of specified runners' names
 *    @param {Array<Array>} runnersPerfs        - Array of arrays containing runners' collated & dated performances
 *    @param {number} [groupPerfsRow=2]         - Row index where to start placing performances on the sheet
 *    @param {number} [groupPerfsCol=4]         - Column index likewise
 *    @param {string} [dateFormat='d-MMM-yy']   - Format used for displayed dates
 *    @param {number} [decimalPlaces=4]         - Floating precision (unless integer)
 * @returns {Range} where performances are on the sheet (beyond subsequent charts)
 */
function CopyRunnersPerformancesToSheet(
  perfsSheet,runners,runnersPerfs,
  groupPerfsRow = 2,
  groupPerfsCol = 4,
  dateFormat = 'dd-MMM-yy',
  decimalPlaces = 4)
{
  var dates = ['Date',...runnersPerfs.map(row =>
    Utilities.formatDate(new Date(row[0]),'UTC',dateFormat))];
  var groupTable = [dates];
  for (var i=1; i<=runners.length; i++) {
    var runnerPerfs = runnersPerfs.map(row => row[i]);
    var runnerName = runners[i-1];
    var perfRow = [runnerName];
    runnerPerfs.forEach(value => {
      var formattedValue;
      if (value === undefined || value === null)
        formattedValue = null;
      else {
        formattedValue = Number(value);
        if (formattedValue % 1 !== 0)
          formattedValue = Number(formattedValue.toFixed(decimalPlaces));
      }
      perfRow.push(formattedValue);
    });
    groupTable.push(perfRow);
  }
  var groupPerfsRange = perfsSheet.getRange(
    groupPerfsRow,groupPerfsCol,
    groupTable.length,groupTable[0].length);
  groupPerfsRange.setValues(groupTable);
  return groupPerfsRange;
}

/**
 * Apply format to group performances and prepare space on sheet for the chart
 *     @param {Sheet}  perfSheet       - The sheet where spacing & formatting applies
 *     @param {number} runnersLegend   - The runners in the legend (to be coloured)
 *     @param {Range}  groupPerfsRange - The range of performances (with Date header)
 *     @param {number} perfChartRow    - The row where the chart will be placed
 *     @param {number} perfChartCol    - The column where the chart will be placed
 *     @param {string} perfFormat      - The format to apply to performance values
 *     @param {number} chartHeight     - The height of the chart
 *     @param {number} chartWidth      - The width of the chart
 *     @param {number} offsetBorder    - The offset border size (in pixels)
 */
function ApplyFormatsOnGroupPerformancesAndChartSheet (
  perfSheet,groupPerfsRange,runnersLegend,
  perfChartRow,perfChartCol,perfFormat,
  chartHeight,chartWidth,offsetBorder=offsetBORDER)
{
  const padFACTOR = 0.05;       // adjust min/max values to encourage "growth"
  const cellHeightSIZE = 21;    // cell height accounts for merged cells on chart
  const surroundFACTOR = 2;     // pad out border all around the chart (in pixels)
  const rotateDownANGLE = -90;  // ensures Dates header text appears downward
  perfSheet.setRowHeight(perfChartRow,chartHeight+surroundFACTOR*offsetBorder
    -cellHeightSIZE*runnersLegend.length);  // Adjust for merging rows (later)
  perfSheet.setColumnWidth(perfChartCol,chartWidth+surroundFACTOR*offsetBorder);
  perfsRange = groupPerfsRange.offset(1,1,  // Get runners' performances only...
    groupPerfsRange.getNumRows()-1,         // ...ignoring the Date row
    groupPerfsRange.getNumColumns()-1);     // ...and the runner name column
  var perfsValues = perfsRange.getValues(); // Values unaffected by formatting
  var minPerf = Math.min(...perfsValues.flat()
    .filter(x => x !== null && x !== undefined && x !== ''));
  minPerf = minPerf*(1-padFACTOR);
  var maxPerf = Math.max(...perfsValues.flat());
  maxPerf = maxPerf*(1+padFACTOR);
  perfsRange.setNumberFormat(perfFormat);   // Apply the specified format
  groupPerfsRange.offset(0,0,1)   // Assume top date row format is ok,,,,
    .setHorizontalAlignment("center")   // ...align cells to the centre, and
    .setVerticalAlignment("middle")     // ...middle (based on a tall row), and
    .setTextRotation(rotateDownANGLE);              // ...rotate text by 90° downwards
  let numGroupPerfsRows = runnersLegend.length+1; // = num<group>PerfsROWS
  perfSheet.getRange(perfChartRow,perfChartCol,   // Merge the rows for the Chart
    numGroupPerfsRows,1).merge();       // ...adjusted to # of runners in the group
  return { min: minPerf, max: maxPerf };
}

/**
 * Embeds a line chart of specified groups of runners' performances in an existing sheet.
 *  Prior to creating a Group chart for comparing related runners, it is assumed that
 *  the performances have already been collated from each of the grouped runners 
 *  results and have already been presented to the right of where this Group chart is
 *  to be created on this same sheet.
 *    @param perfSheet         - The sheet (object) to embed the chart in
 *    @param {string} title    - The title of the chart
 *    @param {range} groupPerfsRange      - Performances range on RHS of chart position
 *    @param {Array<Array<string>>} runnersLegend - A 2D array of runner names & colours.
 *    @param {number} [perfChartRow=2]    - The row number to place the chart at
 *                                           (avoid subsequent charts coinciding)
 *    @param {number} [perfChartCol=2]    - The column number to place the chart 
 *                                           (typically starting at B2)
 *    @param {boolean} [showDates=false]  - Allow non-contiguous dates on the horizontal
 *    @param {boolean} [reverseTrend=1]   - Trends in reverse: rising to low values
 *    @param {number} [filterRecentYears=recentYRS] - The cut-off years for results 
 *    @param {string} [perfTitle=chartYAxisTITLE]   - Performance title on vertical
 *                                                    (typically Age Grade in results)
 *    @param {string} [perfFormat=ageGradeFORMAT]   - Format the performance (e.g. %age)
 */
function EmbedRunnersPerformancesChartInSheet(
  perfSheet,title,groupPerfsRange,runnersLegend,
  perfChartRow = 2,
  perfChartCol = 2,
  showDates = false,reverseTrend = 1,
  filterRecentYears = recentYRS,
  perfTitle = chartYAxisTITLE,
  perfFormat = ageGradeFORMAT)
{
  const chartWIDTH = 800;
  const chartHEIGHT = 350;
  const offsetBORDER = 5; // pixels
  var perfLimits = ApplyFormatsOnGroupPerformancesAndChartSheet(
    perfSheet,groupPerfsRange,runnersLegend,
    perfChartRow,perfChartCol,perfFormat,
    chartHEIGHT,chartWIDTH,offsetBORDER);
  if (filterRecentYears > 0)
    title += ' (max '+filterRecentYears+' years)';
  let colours = runnersLegend.map(colour=>colour[1]);
  var seriesOptions = {};
  for (var i=0; i<colours.length; i++) {
    seriesOptions[i] = { color: colours[i] };
  }
  var embeddedChart = perfSheet.newChart()
    .asLineChart()
    .addRange(groupPerfsRange)
    // .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)  // if multiple sheet ranges?
    // .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH) // always shows
    .setTransposeRowsAndColumns(true)         // data on same sheet in D column?
    .setNumHeaders(1)
    .setPosition(perfChartRow,perfChartCol,offsetBORDER,offsetBORDER)
    .setXAxisTitle('Date')
    .setYAxisTitle(perfTitle)
    .setOption('useFirstColumnAsDomain',true) // Dates on x-axis (in 1st Row before transpose)
    .setOption('curveType','function')
    .setOption('interpolateNulls',true)
    .setOption('legend.position','top')
    .setOption('title',title)
    .setOption('vAxis.direction',reverseTrend)  // Low values at the top (good if it worked!)
    .setOption('vAxis.minValue',perfLimits.min)
    .setOption('vAxis.maxValue',perfLimits.max)
    .setOption('treatLabelsAsText',showDates)   // show all dates on chart
    .setOption('isStacked','false')  // performances not aggregated
    .setOption('width',chartWIDTH)
    .setOption('height',chartHEIGHT)
    .setOption('series',seriesOptions)
    // .setOption('animation.duration',500)
    .build();
  perfSheet.insertChart(embeddedChart); // Manual changes to charts possible thereafter?
  return embeddedChart;
}

/**
 * Retrieves the Performances sheet and clears any charts from it.
 *  A blank sheet is created if the Performances sheet does not exit
 *    @param {string} [perfSheetName="Performances"] - The name of the sheet to retrieve and clear.
 *  @returns {Spreadsheet.Sheet} The cleared sheet object.
 */
function GetClearPerformancesSheet(
  perfSheetName = "Performances")
{
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var perfSheet = spreadsheet.getSheetByName(perfSheetName);
  if (!perfSheet) {
    perfSheet = spreadsheet.insertSheet(perfSheetName);
    if (!perfSheet) return null;
  }
  var charts = perfSheet.getCharts();
  var numCharts = charts.length;
  if (numCharts > 0) {
    charts.forEach(function(chart) {
      perfSheet.removeChart(chart);
    });
  }
  return perfSheet;
}

/**
 * Extracts runner names and corresponding colours from a range.
 *    @param {Range} runnersRange - contains runners' names in 1st & colors in 2nd rows.
 * @returns {Array<Array<string>>} 2D array of [name,color] pairs for each runner.
 */
function ExtractGroupRunners(runnersRange) {
  var runnersNames = runnersRange.getValues()[0];
  var runnersColours = runnersRange.getValues()[1]; // aligned with names above
  var runnersLegend = [];
  for (var j=0; j<runnersNames.length; j++) {
    if (runnersNames[j] != "") {
      runnersLegend.push([runnersNames[j],runnersColours[j]]);
    }
  }
  return runnersLegend;   // 2-D array
}

/**
 * Generate a performance chart in a specified sheet for a group of runners
 *    @param {string} perfSheet     - The name of the sheet to generate the chart in
 *    @param {string} chartTitle    - The full title of the chart
 *    @param {Array<Array<string>>} runnersLegend   - A 2D array runner,colour pair
 *    @param {number} perfChartRow  - The row number to position the chart
 *    @param {number} perfChartCol  - The column number to position the chart.
 *    @param {string} showDates     - Allows for long gaps (false unless true)
 *    @param {string} reverseTrend  - Reverse vertical axis (1, unless -1)
 *    @param {number} filterRecentYears - Cut-off excess years (0 if no filter)
 *    @param {string} perfColumnTitle   - The vertical title (default: Age Grade)
 *    @param {number} perfColumnIndex   - The index of the performance  (5 = F)
 *    @param {string} perfFormat        - The format to apply to performances
 */
function GeneratePerformanceChartInSheet(
  perfSheet,chartTitle,runnersLegend,
  perfChartRow,perfChartCol,
  showDates=false,reverseTrend=1,
  filterRecentYears,
  perfColumnTitle=chartYAxisTITLE,perfColumnIndex,perfFormat)
{
  if (!perfSheet) return null;
  let runners = runnersLegend.map(runner=>runner[0]);
  var runnersDatedPerfs = GetRunnersDatedPerformances(
    runners,filterRecentYears,perfColumnIndex);
  if (!runnersDatedPerfs) return null;
  var groupPerfsRange = CopyRunnersPerformancesToSheet(perfSheet,runners,
    runnersDatedPerfs,perfChartRow);
  if (!groupPerfsRange)
return null;
  var perfChart = EmbedRunnersPerformancesChartInSheet(
    perfSheet,chartTitle,groupPerfsRange,runnersLegend,
    perfChartRow,perfChartCol,
    showDates,reverseTrend,
    filterRecentYears,
    perfColumnTitle,perfFormat);
  if (perfChart) {
    Logger.log('Generated group chart, '+chartTitle+
      ' in row, '+perfChartRow+' of sheet '+perfSheet.getName());
  }
}

/**
 * Generates charts in performance sheet(s) based on config in the Groups sheet.
 *    @param {string} [groupsSheetName="Groups"] - existing sheet with Groups config
 */
function GenerateChartsFromGroupsSheet(
  groupsSheetName = "Groups")
{
  var groupsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(groupsSheetName);
  if (!groupsSheet)
return;
  const numGroupROWS = 3;       // Group of related runners info spread over 3 rows 
  const startGroupROW = 3;      // Asumme title and header in top 2 rows above
  const lastGroupROW = groupsSheet.getLastRow();
  const maxGroupsCOUNT = parseInt((lastGroupROW-startGroupROW+1)/numGroupROWS);
  const startRunnerCOL = 3;     // Assume list of runners start in column C
  const maxRunnersCOUNT = groupsSheet.getLastColumn()-startRunnerCOL+1;
  const paramsCOUNT = 9;        // Assume parameters of Group in Columns A..I
  // For each Group, there are THREE ordered steps to generate the chart...
  var defaultPerfSheet = null;  // Assume default Performance Seeet unless specified
  var is1stGroup = true;
  for (var i=0; i<maxGroupsCOUNT; i++) {
    var group1stRow = startGroupROW+numGroupROWS*i; // Assumes no blank rows between
    // 1. Establish valid Group in runners (2D-Array) with colours on two rows...
    var runnersRange = groupsSheet.getRange(group1stRow+1,  // Runner names on 2nd row
      startRunnerCOL,2,maxRunnersCOUNT);                // with legend on next row
    var runnersLegend = ExtractGroupRunners(runnersRange);
    if (!runnersLegend || runnersLegend.length == 0)
  continue; 
    // 2. Extract Group parameters from 1st row, except chart position from next row
    //    Assume null for function defaults (if empty) - see notes on Groups sheet
    var paramsRange = groupsSheet.getRange(group1stRow,1,1,paramsCOUNT);
    var perfSheet,perfSheetName,groupName,groupTitle,
      showDates,reverseTrend,
      filterRecentYears,
      perfColumnTitle,perfColumnIndex,perfFormat;
    var params = paramsRange.getValues()[0];
    for (var j=0; j<params.length; j++) {   // ensure all parameters considered
      switch (j) {
      case 0:     // column A (of Groups sheet)
        perfSheetName = params[0] || defaultPerfSheet;
        if (is1stGroup || perfSheetName != defaultPerfSheet) {
          perfSheet = GetClearPerformancesSheet(perfSheetName);
          defaultPerfSheet = perfSheetName;   // if explicit assume default thereafter
          is1stGroup = false;
        } else {
          perfSheet = SpreadsheetApp.getActiveSpreadsheet()
            .getSheetByName(perfSheetName);
        }                                             break;
      case 1:     // column B
        groupName = params[1];                        break;
      case 2:     // column C
        groupTitle = params[2];                       break;
      case 3:     // column D
        // used for vertical & main titles , and
        // ... matches a header in results sheets
        perfColumnTitle = params[3] || chartYAxisTITLE;  break;
      case 4:     // column E
        perfColumnIndex = params[4] || undefined;     break;
      case 5:     // column F
        perfFormat = params[5] || undefined;          break;
      case 6:     // column G
        filterRecentYears = params[6] || undefined;   break;
      case 7:     // column H
        showDates = params[7] == true || false;     break;
      case 8:     // column I
        reverseTrend = params[8] == true ? -1 : 1;  break;
      }
    } // end of parameters for one Group from Groups sheet
    var perfChartRow = paramsRange.offset(1,0)
      .getValue() || undefined;  // in column A on next row
    var perfChartCol = paramsRange.offset(1,1)
      .getValue() || undefined;  // in column B
    if (!groupName)
  continue
    // 3. Place Group performances data on sheet before creating Group chart
    var chartTitle = groupName+" ("+groupTitle+") "+perfColumnTitle;
    GeneratePerformanceChartInSheet(
      perfSheet,chartTitle,runnersLegend,
      perfChartRow,perfChartCol,
      showDates,reverseTrend,
      filterRecentYears,
      perfColumnTitle,perfColumnIndex,perfFormat);
  } // end of Groups from Groups sheet
  return perfSheet;
}

function SetLegendCellColoursOnSheet(sheetName="Groups") {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  for (var i=0; i<values.length; i++) {
    if (values[i][1] === "Legend") {
      for (var j=1; j<values[i].length; j++) {
        var hexValue = values[i][j];
        if (typeof hexValue === "string"
            && hexValue.match(/^#[0-9A-F]{6}$/i))
        {
          sheet.getRange(i+1,j+1).setBackground(hexValue);
          sheet.getRange(i+1,j+1).setFontColor("#ffffff");
        }
      }
    }
  }
}

/**
 * Sets up the Parkruns menu on opening the spreadsheet.
 *  Instructions to set up (and execute?) the trigger:
 *    1. Go to the Apps Script (Macros) editor.
 *    2. Click on the clock icon (Triggers) in the left sidebar.
 *    3. Click on "Create trigger".
 *    4. Set up the trigger with the following settings:
 *        - Choose function: onOpen
 *        - Select event type: On open
 *        - Save
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var parkrunsMenu = ui.createMenu('Parkruns')
    .addItem("Import lastest result for each runner"+
      "\u00A0".repeat(7)+"Ctrl+Alt+Shift+9",'ImportLatestResultForEachRunner')
    .addItem("Clean format for multiple results"+
      "\u00A0".repeat(13)+"Ctrl+Alt+Shift+1",'CleanFormatForPastedRunResults')
    .addItem("Re-protect results for each runner"+
      "\u00A0".repeat(11)+"Ctrl+Alt+Shift+7",'ReprotectEachRunnerResultsSheets')
    .addItem("Generate charts from Groups"+
      "\u00A0".repeat(19)+"Ctrl+Alt+Shift+0",'GenerateChartsFromGroupsSheet')
    .addItem("Colour legends for Groups",
      'SetLegendCellColoursOnSheet')
    // .insertMenu(ui,5)   // ideally before Tools 
    .addToUi();
}

function PasteAboveRangeFormula() {
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  var activeCells = sheet.getActiveRange();
  var aboveCells = sheet.getRange(
    activeCells.getRow()-1, 
    activeCells.getColumn(), 
    activeCells.getNumRows(), 
    activeCells.getNumColumns()
  );
  aboveCells.copyTo(activeCells, SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
}

function DuplicateAboveRowFormula() {
  var sheet = SpreadsheetApp.getActive().getActiveSheet();
  var activeRange = sheet.getActiveRange();
  var row = activeRange.getRow();
  sheet.insertRowBefore(row);
  var aboveRow = sheet.getRange(row,1,1,sheet.getLastColumn());
  var originalRow = sheet.getRange(row+1,1,1,sheet.getLastColumn());
  originalRow.copyTo(aboveRow,SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
}

function swapColumns() {
  // Swap columns selected by user (potentially across multiple sheets)
  var selection = SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList().getRanges();
  // Check if selection is valid (2 columns, same sheet or extended selection)
  if (selection.length < 2) {
    SpreadsheetApp.getUi().alert('Please select at least two columns to swap.');
    return;
  }
  var columns = selection.map(range => range.getColumn());
  if (new Set(columns).size !== 2) {
    SpreadsheetApp.getUi().alert('Please use Ctrl+ to select two columns explicitly.');
    return;
  }
  // Get unique sheets from selection
  var sheets = new Set(selection.map(range => range.getSheet()));
  // Swap columns on each sheet
  sheets.forEach(function(sheet) {
    var lastRow = sheet.getMaxRows();
    var range1 = sheet.getRange(1, Math.min(...columns), lastRow);
    var range2 = sheet.getRange(1, Math.max(...columns), lastRow);
    var temp = range1.getValues();
    range1.setValues(range2.getValues());
    range2.setValues(temp);
  });
  // SpreadsheetApp.getActiveSpreadsheet().getActiveRangeList().removeAllRanges();
}
