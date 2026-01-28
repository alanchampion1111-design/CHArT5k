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
/   4.  Import result for each runner
/       (ImportResultForEachRunner - Ctrl+Alt+Shift+9)
/   5.  Generate charts from Groups
/       (GenerateChartsFromGroups - Ctrl+Alt+Shift+2)
/ 	6.  Colour legends in Groups
/       (ColourLegendsInGroups)
/---------------------------------------------------------------------------------------
 */

/**
 * @OnlyCurrentDoc
 *  // Ensure authorisation granted via appscript.json
 * @scope https://www.googleapis.com/auth/script.external_request
 * @scope https://www.googleapis.com/auth/script.scriptapp
 * @scope https://www.googleapis.com/auth/spreadsheets
 * @scope https://www.googleapis.com/auth/script.container.ui
 * @scope https://www.googleapis.com/auth/drive.readonly
 * @scope https://www.googleapis.com/auth/drive
 */

const resultTABLE = "Event"   // For any Runner Event results sheet
const eventHeaderCELL = "A2"; //  where the header row is below the Runner name title
const firstResultCELL = "A3"; //  and at least one Result has been cleanly entered
const firstResultROW = 3;     //  where new Results MUST be below this first result
const scrollCOL = 13;      // Scroll down when selecting column M beyond the table
const pasteCOL = 1;        // Paste Event result starting in 1st column (A)
const timeCOL = 5;         // Event result time is in the 5th column (E)

// 
const runnersSheetNAME = 'Runners';
const allRunnersSHEET =           // The Runners sheet drives the creation of results sheets
  SpreadsheetApp.getActiveSpreadsheet()   // ...for new runners by those permitted
    .getSheetByName(runnersSheetNAME);
const templateNameCELL = "J1";          // This cell identifes the seed template (e.g. Keren)
const templateNAME = allRunnersSHEET    // The seed template may be readily reconfigured
  .getRange(templateNameCELL)     // See the note on Runners!J1 cell
  .getValue();                    //  ...where only the first 3 (or 4) rows are relevant
const parkrunnerIdCOL = 10;       // in column J on Runners sheet
const parkrunnerIdINDEX = 9;      // in column J on Runners sheet (for arrays or range offsets)
const runnersStartROW = 3;        // start after title & header rows (2)
const resultsStartROW = 3;      // start after title & header rows (2)
const numBlankROWS = 5;          // Regular catch-up of multiple results by those permitted
const dateFORMAT = 'd-MMM-yy';    //consistent for backwards compatibility

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

/**
 * Ensures a specified number of blank rows exist at the end of the active sheet.
 *  If necessary, inserts new rows to accommodate the specified number.
 *    @param {number} numRows - The number of blank rows to ensure (default: 10)
 *  @return {string} The A1 notation of the range of blank rows
 */
function EnsureBlankResultsRange(numRows=numBlankROWS) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var runnerNameId = sheet.getName();
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
  logMessage += " on results sheet, "+runnerNameId;
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
  var runnerNameId = sheet.getName();
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
  var addersCell = allRunnersSHEET.getRange(addersCOLUMN+(indexRow+runnersStartROW));
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
 * Reprotects each runner's results sheet, to allow permitted users to add results.
 *  For each named runner...
 *    0. Get that runner's results sheet (or create if none)
 *    1. Ensure sufficient blank rows to add new result(s)
 *    2. Protect that results sheet except for the blank row range
 *    3. Protect that blank range for specific (editor) runners only to add results
 */
function ReprotectEachRunnerResultsSheets() {
  const runnersRANGE = runnerNameCOLUMN+resultsStartROW+":"+runnerSurnameCOLUMN;
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
    var runnerNameId = runnerName+'_'+index;
    var resultsSheet = spreadsheet.getSheetByName(runnerNameId);
    if (!resultsSheet) {
      var thisRow = index+runnersStartROW;
      var parkrunnerId = allRunnersSHEET.getRange(
        thisRow,parkrunnerIdCOL).getValue();
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
/   The functions involved are significantly disjointed, relying on manual steps,
/   and reliant on permission being granted beforehand
/
/     Loop manually for each runner...
/       Copy the latest result rows (if missing set)...
/         onSelectionChange (by pressing down arrow icon)
/           ScrollBeyondLastResult (may create rows if permitted)
/         Paste manually into the runner's result sheet...
/           CleanFormatforPastedRunResults (after copying row from parkrun event)
/             ClearBordersOnRange
/             ApplyFormatFromRowAboveOnRange
/             ApplyFormulaFromRowAboveBeyondRange
/             RepairDurationTimesInRangeColumn
/               TranslateDurationFromHoursToMinutes
/               
/  ----------------------------------------------------------------------------
*/

function ScrollBeyondLastResult() {
  const functionNAME = "ScrollBeyondLastResult";
  // Select the first empty row in preparation for pasting a non-parkrun entry
  // Conveniently used to skip to the bottom of the table below the latest result
  var sheet = SpreadsheetApp.getActiveSheet();
  Logger.log('Running '+functionNAME+' on sheet: '+sheet.getName());
  var lastRow = sheet.getLastRow();
  if (lastRow <= firstResultROW) return; // table has no results yet
  // Find an empty row from the bottom of the table since more efficient
  var i = lastRow;
  for (; i>firstResultROW; i--) 
    if (sheet.getRange(i,pasteCOL).getValue()) break;
  // When last row has a result, prepare for pasting new result(s) below
  if (i == lastRow)       
    sheet.appendRow([""]);
  // Always select the row below the last result
  sheet.getRange(i+1,pasteCOL).activate();
}

/**
 * Sets up skipping to the bottom (on selecting M3 cell)
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
  if (range.getColumn() === scrollCOL)
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
 *    @param {number} [timeCol=timeCOL] - Column containing the duration times
 */
function RepairDurationTimesInRangeColumn(range,
  timeCol = timeCOL)
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
    var pastedValue = sheet.getRange(firstRow,pasteCOL).getValue();  
    if (!pastedValue) {
      SpreadsheetApp.getUi().alert('Error','Ensure that the result(s) have been copied from the All Results table'+
        '(with the Event Name and PB?) and that the pasted result(s) to be cleaned remain selected',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    var pastedColumn = pastedResults.getColumn();
    if (pastedColumn !== pasteCOL) {
      SpreadsheetApp.getUi().alert('Error','Ensure that the result(s) have been pasted into the first column (A) and that they remain selected',
        SpreadsheetApp.getUi().ButtonSet.OK);
      return;   
    }
    var firstTimeString = sheet.getRange(firstRow,timeCOL).getDisplayValue();
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
  RepairDurationTimesInRangeColumn(pastedResults,timeCOL);
}

/* --------------------------------------------------------------------------
/
/   The following definitions and functions automate the addition of a new
/   result by securely connecting to the parkrun site for importing the
/   latest result, mimicking a manual extract to extract copy and paste.
/
/   There are two use-cases, where the heirarchy of function calls are:
/
/   1.  ImportResultForEachRunner (optional event date of Parkrun)
/         OpenChromeBrowser ->setValues
//        Loop for each member runner...
/           >CopyResultForRunner (latest or dated?)->
/             >GetRunnerResultsPage
/               >AccessPage
/             GetResultRow
//          When latest is a new result...
/             >PasteResultForRunner ->
//              Loop for each cell...
/                 CleanValue (apply links as hyperlinks)
/               FormatDate
/               AppendResultRow
/             >AppendPositionsForResult ->
/               GetResultUrl (assume always parkrun)
/               GetDomainGender  (for language-based filter)
/               ExtendRange
/               >AssessPositions (for one runner)
/               IncludePositions
/         >CloseChromeBrowser
/
/   2.  CatchupAllPositions
/         OpenChromeBrowser ->
//        Loop for each member runner...
/           LockCallerForwardsTo ->
/             BatchPositionsForRunner (threaded in parallel)
/               UnlockCallerForwarded
/               SyncPositionsPerRunner -> (in batches of 10 results)
/                 FirstMatchRange (based on unknown Gender position)
//                When positions not previously updated...
//                  Loop for each runner's result (per batch)...
/                     AppendPositionsForResult ->
/                       GetResultUrl
//                      When event is a Parkrun...
/                         GetDomainGender (for language-based filter)
/                         >AssessPositions (for one runner)
/                         IncludePositions
/               >BatchPositionsForRunner (recurse if more to sync)
/               : (nested executions possible)
/                 >CloseChromeBrowser
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
// const sampleURL = 'https://www.example.com';  // default test
const sampleURL = 'https://www.parkrun.org.uk/colchestercastle/results/116';
const debug = true;   // WARNING: debug if true may slow down performance and may skip runners!!
var browserSession;   // Shared by threaded/recursed processes for up to 30 minutes because browser lingers

async function OpenChromeBrowser() {
  const initBrowserURL = browserURL+'/initBrowser';
  try {
    var response = await UrlFetchApp.fetch(initBrowserURL);
    browserSession = response.getContentText();   // WS Endpoint lingers for upto 30 minutes
    Utilities.sleep(1000); // in case image not cached, avoids Error: Requesting main frame too early!
    Logger.log('Session: '+browserSession);
  } catch (err) {
    Logger.log(err);
  }
}

async function AccessPage(
  thisUrl = sampleURL)
{
  let getBrowserURL = browserURL+'/getUrl?url='+thisUrl;
  try {
    var response = await UrlFetchApp.fetch(getBrowserURL,
      // Runners with 500+ results take longer 
      //    AND parallelism may be curtailed
      //      WHEN re-using the same page is detected,
      //      AND/OR WHEN a single CPU is specified for the run
      // Therefore, allow worst case timing for all, as if serial
      {muteHttpExceptions:true,timeout:30000});
    Utilities.sleep(1000); // in case client-side table needs refresh (aside awaiting the fetch)
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
  	Logger.log('Session ended. '+response.getContentText());
    browserSession = undefined;
  } catch (err) {
    Logger.log(err);
  }
}

// Although driven from UK site, specific results are on different domains in other countries
const parkrunURL = 'https://www.parkrun.org.uk/';
const parkrunnerURL = parkrunURL+'parkrunner/';
const allForONE = '/all/';
const oneForALL = '/results/';    // TODO: Potentially cache if same location as next member
var cacheResultsURL = undefined;  // ...assumes cached within the browser service for re-use

async function GetRunnerResultsPage(
  parkrunnerId = '1213963')    // default useful for testing
{
  let thisParkrunnerURL = parkrunnerURL+parkrunnerId +allForONE;
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
 *  Gets the latest result or a dated result if a date is specified (dd/mm/yyyy)
 * 
 */
/**
 * Gets a specific result row for a runner (default to latest in the first row)
 *  @param {Array<string>} bodyRows - Array of HTML rows containing result data.
 *  @param {number|string} [eventRef=0] - Index of the result (0 = latest) or date (dd/mm/yyyy) of result.
 * @returns {string|null} The matching result row HTML, or null if a matching date is not found
 * Assumes operates within "universal" (non-US) date browser setting.
 */
function GetResultRow(
  bodyRows,
  eventRef = 0)  // default to latest (indexed 0) in first row; otherwise match as a date 
{
  const latestROW = 0;  // only consider 1st row as single latest result
  if (typeof eventRef === 'number') {
    return bodyRows[eventRef];
  } else {
    return bodyRows.findIndex(row => row.includes(`>${eventRef}<`))
  }
}

/**
 * Copies the latest result for a given parkrunner ID from Parkrun site
 *    @param {string} parkrunnerId - The ID of the parkrunner (default: "21283")
 *    @eventRef (string) - ALL or date (dd/mm/yyyy assumes UK format setting in the browser)
 *    @thisPage (html) - typically preloaded when ALL only
 *  @returns {array|null} - represents the latest result (or null if none)
 */
async function CopyResultForRunner(
  parkrunnerId = "777764",
  eventRef = undefined,  // by default.the latest; otherwise find one that matches dd/mm/yyy
  thisPage = undefined)   // pre-loaded Resuts Page 
{
  const allResultsTITLE = "All  Results";    // 3rd table with All Results
  var headerStart = "<thead><tr><th>Event</th><th>Run Date</th><th>Run Number</th><th>Pos</th>";
  const allHEADER = headerStart+"<th>Time</th><th>Age<br/>Grade</th><th>PB?</th></tr></thead>";
  // WARNING: Row does not (yet) show Gender position,
  //          nor positions for Age-Category or Age-Grade (%age)
  return GetRunnerResultsPage(parkrunnerId)
  .then (allResults => {   // ensures allResults is not a Promise
    if (allResults) {
      allResults = allResults
        .substring(allResults
          .indexOf(allResultsTITLE))  // trim before "All  Results"
        .split('#comments')[0];      // trim after #comments
      if (allResults) {
        // if (debug) Logger.log('All Results table: '+allResults); 
        var headerRow = allResults.match(/<thead>.*?<\/thead>/s);
        if (headerRow && headerRow.length>0) {  
          // if (debug) Logger.log('Header row: '+headerRow[0]);  // TODO: columns may change - WARN if so?
          var bodyContent = allResults.match(/<tbody>.*?<\/tbody>/s)[0];
          var bodyRows = bodyContent.match(/<tr[^>]*>.*?<\/tr>/gs);
          bodyRows = bodyRows.filter(row => !row.includes('junior'));    // junior 2km times upset rankings
          if (bodyRows && bodyRows.length>0) {
            if (eventRef === 'ALL') {   // return 2D array of results
              return bodyRows.map(row => {
                var cells = row.match(/<td>.*?<\/td>/gs);
                return cells.map(cell => cell.replace(/<td>|<\/td>/g, "").trim());
              });
            } else { // // single result, default to latest
              var resultRow = GetResultRow(bodyRows,eventRef);  
              if (resultRow) {
                var cells = resultRow.match(/<td>.*?<\/td>/gs);
                if (cells) {
                  var values = cells.map(function(cell) {
                    return cell.replace(/<td>|<\/td>/g,"").trim();
                  });
                  // if (debug) Logger.log('Cells parsed as '+cells+' for runner, '+parkrunnerId+', and ready for pasting');
                  return values;
                } else {
                  Logger.log('WARNING: No cells found in row, '+resultROW+' for runner, '+parkrunnerId);
                  return null;
                }
              } else {
                Logger.log('Unable to find results for runner,'+parkrunnerId);
                return null;
              }
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


let hyperLinks = [];  // Preserves the links from the runner's page
const PBtickBoxCOL = 8; // column H is a read-only tick-box, derived from column G (if PB present)

/**
 *  Clean (each) cell, perhaps double clean if necessary)
 *  ...while formatting the values as hyperlinks where so on the page
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
 *    @param {string} dateFormat - Output date format (e.g. "d-MMM-yy")
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
 *  Appends a result row (columns A-G) into results sheet (including 2 or more hyperlinks)
 *    @param {string[]} thisResult - Latest result (cell values) for a runner
 *    @param {Sheet} resultsSheet - Sheet to append the latest result row
 *    @param {offset} default append (or update last if first row)
 *  @returns {range} resultsRow - contains the results link to determine detailed positions
*/
function AppendResultRow(
  thisResult,   // cols A..G (7)
  resultsSheet)
{
  try {
    const linksNUM = hyperLinks.length;             // cols A..C (3)
    const valsNUM = thisResult.length-linksNUM;     // cols D..G (4)
    const dateLinkCOL = 2;  // Col B
    thisResult = thisResult.map(cellValue => {
      if (hyperLinks.length) {    // initially, assume 3?
        return '=HYPERLINK("'+
          hyperLinks.shift()      // ..with consuming side-effect
          +'","'+cellValue+'")';
      } else {        // D..G as normal
        return cellValue;
      }
    });
    // Apply likewise differently to the linked cells
    var insertRow = resultsSheet.getLastRow()+1;  // Assumes MAP formula in 1st Row is never the last
    resultsSheet.getRange(insertRow,1,1,linksNUM)
      .setFormulas([thisResult.slice(0,linksNUM)]);   // A3:C3 first row
    resultsSheet.getRange(insertRow,linksNUM+1,1,valsNUM)
      .setValues([thisResult.slice(linksNUM)]);       // D3:G3 first row    
    var pastedRowRange = resultsSheet.getRange(insertRow,1,1,thisResult.length);
    var resultsLink = pastedRowRange
      .getCell(1,dateLinkCOL).getRichTextValue().getLinkUrl();
    Logger.log('Link to detailed results (deferred for batching later):\n'+resultsLink);
    if (insertRow > resultsStartROW)
      resultsSheet.getRange(insertRow,PBtickBoxCOL)   // column beyond paste
        .clear({contentsOnly: true,
          skipFilteredRows: true})    // complement to MAP prevents FALSE 
        .insertCheckboxes();          // c/fwd tickbox restored
    return pastedRowRange;
  } catch (err) {
    Logger.log('ERROR: Unable to add hyperlinks correctly to latest result for runner...\n'+err);
    return null;
  }
}

/**
 * Pastes the latest result for a given runner into their sheet.
 *    @param {array} thisResult - representing the latest (single) result
 *    @param {string} [runnerName="Alan"] - name of the runner's results sheet
 *  @returns {Range} - where result was pasted (or null if result already pasted)
 */
 function PasteResultForRunner(
  thisResult,         // assumed a single row to be added
  runnerNameId = "Alan_13")
{
  if (debug) Logger.log('Consider pasting result into unique runner sheet, '+runnerNameId+'...');
  const locationINDEX = 0;  // column A
  const dateINDEX = 1;      // column B
  const previousROWS = 15;  // previous results may include non-parkrunner events
  hyperLinks = [];   // discard any hyperLinks from duplicate results
  thisResult = thisResult.map(CleanValue);   // includes 00: (for hh:) prefix on elapsed time
  var thisDate = thisResult[dateINDEX] = FormatDate(thisResult[dateINDEX],dateFORMAT);
  var thisLocation = thisResult[locationINDEX];
  var resultsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(runnerNameId);
  // Consider matching only with recent previous results (after ignoring title & header rows)
  // This takes account of some runners who run elsewhere (and captured manually) since their last recorded parkrun
  var previousResults = resultsSheet.getDataRange().getDisplayValues().slice(2).slice(-previousROWS);
  if (previousResults.length < 1) {
    Logger.log('ERROR: No previous result for unique runner, '+runnerNameId+' (with map formulae?) to be able to match new result');
    return null;
  }
  var matchingResults = previousResults.map(function(row) {
    return row[locationINDEX]+'&'+row[dateINDEX];
  });
  if (matchingResults.includes(thisLocation+'&'+thisDate)) {    // check duplicate?
    hyperLinks = [];  // discard pending hyperlinks as with this duplicate result
    if (debug) Logger.log('Previously captured result at '+thisLocation+' on '+thisDate+' for unique runner, '+runnerNameId);
    return null;    // null range if not a new result
  } else {
    var pastedResult = AppendResultRow(thisResult,resultsSheet);
    Logger.log('New result at '+thisLocation+' on '+thisDate+' added for unique runner, '+runnerNameId);
    return pastedResult;
  }       
}

/**
 * Determines extra positions for a runner from parkrun results page
 *  @param {string} matchRunner
 *  @param {string} thisUrl 
 *  @param {string} ageCat 
 * @return {Promise<Array>} [acPosn,agPosn,gcPosn]
 */
async function AssessPositions(matchRunner,thisUrl,ageCat,genderCat) {
  let filterBrowserURL = browserURL+'/filterUrl'
    +'?url='+thisUrl
    +'&rn='+encodeURIComponent(matchRunner)
    +'&ac='+ageCat
    +'&gc='+genderCat;
    // +'&ag='Age-Grade';   // assume internal default
  try {
    if (thisUrl === cacheResultsURL) {
      if (debug) Logger.log('Runner: '+matchRunner+'\tImplicitly re-using results from previous runner?');
    } else {
      cacheResultsURL = thisUrl;
    }
    var positions = await UrlFetchApp.fetch(filterBrowserURL,
      // Runners with 500+ results take longer 
      //    AND parallelism may be curtailed
      //      WHEN re-using the same page is detected,
      //      AND/OR WHEN a single CPU is specified for the run
      // Therefore, allow worst case timing for all, as if serial
      {muteHttpExceptions:true,timeout:12000});
    let posnText = positions.getContentText();
    if (debug) Logger.log('Runner: '+matchRunner+'\t'+posnText);
    if (posnText.includes('Internal Server Error')) {
      browserSession = undefined; // restart the browser!
      return CloseChromeBrowser()
      .then(() => {
        OpenChromeBrowser()
      });
      return null;
    }

    if (posnText.includes('ERROR')) {
      Logger.log('WARNING: Unmatched '+matchRunner+' in event link, '+thisUrl);
      return null;
    } else {
      var posns = JSON.parse(positions);
      return [posns.acPosition,posns.agPosition,posns.gcPosition];
    }
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
  // WARNING: resultRange normally a sheetrange to update cell values
  const GenderPosnCOL = 9;      // column I (new column extension)
  const AgeCatPosnCOL = 10;     // column J (previously done manually)
  const AgeGradePosnCOL = 11;   // column K (new column extension)
  resultRange.getCell(1,AgeCatPosnCOL).setValue(acPosn);    // col J for Age-Cat position
  resultRange.getCell(1,AgeGradePosnCOL).setValue(agPosn);  // col K for Age-Grade position
  // Pending outcome of proposal to Parkrun site: "Replace the Run # with Gender Position?
  // Alternatively, the number of runners in the event may replace the redundant run number
  resultRange.getCell(1,GenderPosnCOL).setValue(gcPosn);    // col I for Gender position
  // return resultRange;   // Values change, although previously extended range itself is unchanged
}

/**
 * Extends the runner's result range passively, typically one result in a single row
 *  @param {Range} resultRange 
 *  @param {numeric} numRows 
 *  @param {string} agPosn
 */
function ExtendRange(resultRange,numRows,numCols) {
  const AgeCatCOL = 12;   // column L is the age-category, derived from event date  (without conflict)
  // Two derived values are generated from a MAP function set for the first result in the table 
  resultRange = resultRange.offset(0,0,resultRange.getNumRows()+numRows,resultRange.getNumColumns()+numCols);
  // resultRange.getCell(resultRange.getNumRows(),AgeCatCOL)
  //  .setValue(null);  // no conflict in derivation from event date in col B and runner's DoB
  if (debug) Logger.log('Extended range of results by '+numRows+' rows and '+numCols+' columns');
  // resultRange implicitly includes  derived Age-Category in col L (assumes MAP function in 1st result)
  let ageCatCell = resultRange.getCell(1,AgeCatCOL);
  let ageCatCellRef = ageCatCell.getA1Notation();
  let ageCat = ageCatCell.getValue();
  if (debug) Logger.log('Age-Category is '+ageCat+' in new runner row at cell, '+ageCatCellRef);
  return resultRange;
}

/**
 * Gets the URL for the event from the hyperlink (against the date in Col B)
 * Otheresie, legacy solution uses event data and instance from neighbouring cells 
 * Handles specific non-UK event locations by mapping them to their respective parkrun domains.
 * Handles anomalies on some Parkrun Parks including/excluding 'Park' in parkrun domains
 */
function GetResultsUrl(eventDateCell,eventLocation,eventInstance) {
  var resultsLink;
  const eventDate = eventDateCell.getDisplayValue;
  eventDateCell.activate();   // hyperlink may surface when cell selected?Copy
  let richText = eventDateCell.getRichTextValue();
  if (richText) {   // WARNING: Only works for cells with HYPERLINK formaulae
    resultsLink = richText.getLinkUrl();
  } else {   // Legacy complication for past results, the embedded URL was not retrievable!!
     const locationUrlMap = {
      'Nidda': 'com.de',
      'Kagerzoom': 'co.nl',
      'Kralingse Bos': 'co.nl',
      'Delftse Hout': 'co.nl',
      'Zuiderpark': 'co.nl'
    };
    // const location = 'Nidda';
    // const response = await fetch(`https://maps.googleapis.com/maps/api/geocode/json?address=${location}&key=YOUR_API_KEY`);
    // const data = await response.json();
    // const countryCode = data.results[0].address_components.filter(c => c.types.includes('country')).short_name;
    const eventMap = {
      'Bushy Park': 'Bushy',
      'Richmond Park': 'Richmond',
      'Crane Park': 'Crane',
      'Walsall Arboretum': 'Walsall',
      'Wimbledon Common': 'Wimbledon',
      'Medina I.O.W.': 'Medina'
    };
    eventLocation = eventMap[eventLocation] || eventLocation;
    let parkrunDomain = locationUrlMap[eventLocation]
      ? parkrunURL.replace('org\.uk',locationUrlMap[eventLocation])
      : parkrunURL;
    eventLocation = eventLocation.toLowerCase()
      .replace(/[-\.,\'’ ]/g,"");
    resultsLink = parkrunDomain+eventLocation+oneForALL+eventInstance;
    //if (debug) Logger.log('WARNING: Reconstructed URL for event on '+eventDate+' since irretrievable from the Date cell');
  }
  resultsLink = (resultsLink.includes('parkrun')) ? resultsLink : null;
  return resultsLink;
}

/**
 * Returns the gender word for a given Country domain based on 2nd char of Age Category as the code
 *    @param {string} [thisUrl=parkrunURL] - URL to determine domain
 *    @param {string} [gender='M'] - Gender code (M/W)
 *  @returns {string} language translation for Gender code (e.g. "Male" for M, "Vrouw" for W, etc.)
 */
function GetDomainGender(
  thisUrl = parkrunURL,
  thisGender = 'M')
{
  const languageMAP = {
    // assume derived from age categorym: e.g. VW35-39 => W
    'en': { M: 'Male',      W: 'Female' },
    'nl': { M: 'Man',       W: 'Vrouw' },
    'de': { M: 'Männlich',  W: 'Weiblich' },
    'it': { M: 'Uomo',      W: 'Donna' }
  };
  const countryMAP = {
    'uk': 'en',
    'ie': 'en',
    'us': 'en',
    'au': 'en',
    'at': 'de'
  };
  let thisDomain = thisUrl.split('/')[2];
  let thisCountry = thisDomain.split('.').pop();
  let thisLang = countryMAP[thisCountry] || thisCountry;
  return languageMAP[thisLang][thisGender];
}

/**
 * Appends detailed positions to an individual runner's sheet beyond the current range.
 * Used initially where region needs extended OR later on catch up when already so.
 *    @param {string} runnerFullName - Runner's full name
 *    @param {Range} resultRange - Range of result data
 *  @returns {boolean} Success flag
 */
async function AppendPositionsForResult(runnerFullName,resultRange) {
  const eventCOL = 1;           // column A is the event location
  const dateCOL = 2;            // column B is the date of event
  const runNumCOL = 3;          // column C is the instance # at event
  let runNumber = resultRange.getCell(1,runNumCOL).getValue();
  if (!runNumber) return false;   // Skipping any non-Parkrun instance since no positions
  else {
    const genderPosnCOL = 9;      // column I is the Gender position (if present already done)
    const ageCatPosCOL = 10;      // column J is the Age-Category position
    const ageGradePosCOL = 11;    // column K is the Age-Grade (%age) position
    const ageCatCOL = 12;         // column L is the Age Category on the date (derived from DoB)
    const PBtoAgeCatNumCOLS = 5;  // cols H..L for I..K positions between derived cols, H & L
    let eventLocation = resultRange.getCell(1,eventCOL).getValue();
    let eventDateCell = resultRange.getCell(1,dateCOL);
    let extResultRange = (resultRange.getNumColumns() < ageCatCOL) 
      ? ExtendRange(resultRange,0,PBtoAgeCatNumCOLS) // extends to include H..L (for Import use-case)
      : resultRange;             // Result row range previously extended (for Catch-up use-case)
    let ageCategory = extResultRange.getCell(1,ageCatCOL).getValue();  // derived from Date - DoB
    let genderPositionKnown = extResultRange.getCell(1,genderPosnCOL).getValue();  
    if (genderPositionKnown) return false;  // Skipping since extra position(s) already on the sheet
    else {
      try {
        var resultsLink = GetResultsUrl(eventDateCell,eventLocation,runNumber); // if workaround needed
        // var resultsLink = GetResultsUrl(eventDateCell); // behind Date (in col B)
        var genderCategory = GetDomainGender(resultsLink,ageCategory[1]); // e.g. JM10 => M => Male in uk
        if (debug) Logger.log('Link to '+eventLocation+' Parkrun results:\n'+resultsLink);
        let extraPosns = await AssessPositions(runnerFullName,resultsLink,ageCategory,genderCategory);
        if (extraPosns && extraPosns.length === 3) {
          IncludePositions(extResultRange,...extraPosns);  // into cells, I..K on result row
          return true;
        } else {
          Logger.log('ERROR: Unable to assess positions for result at '+eventLocation+
            ' for runner, '+runnerFullName+' ('+resultsLink+')');
          return false;
        }
      } catch (err) {
        Logger.log('ERROR: While appending positions for runner, '+runnerFullName+': '+err);
        return false; // or throw error
      } finally {
        Utilities.sleep(2000);
      }
    }
  }
}

/**
 *  Imports the latest results for each of member runners from Parkrun site,
 *  potentially on a specific date (if scheduled or missed), otherwise the latest is assume
 */
function ImportResultForEachRunner(
  // potentially import missing results for a date = e.g. '27/12/2025' or '01/01/2026'
  eventDate = undefined)  // undefined means latest date - return to this state otherwise
{
  return OpenChromeBrowser().then(() => {   // Browser always launched beforehand...
    var runners = allRunnersSHEET.getRange(
      runnerNameCOLUMN+runnersStartROW+":"+runnerSurnameCOLUMN
    ).getValues().filter(String);
    // Process each runner in parallel BEFORE closing after ALL runners done!
    return Promise.all(runners.map(function(runner,index) {
      let runnerName = runner[0];     // col A
      let runnerNameId = runnerName+'_'+index;
      let parkrunnerId = allRunnersSHEET.getRange(
        runnersStartROW+index,parkrunnerIdCOL).getValue();
      if (debug) Logger.log('Parkrunner ID: '+parkrunnerId);
      if (isNaN(parkrunnerId)) {
        Logger.log('WARNING: Skipping invalid parkrunner Id, '+parkrunnerId);
        return Promise.resolve();
      }
      // Ensure ALL functions complete by returning the Promise chain
      return CopyResultForRunner(parkrunnerId,eventDate)
      .then(thisResult => {
        if (thisResult) {
          let resultRange = PasteResultForRunner(thisResult,runnerNameId);
          if (resultRange) {
            if (debug) Logger.log('Appending result for unique runner sheet, '+runnerNameId);
            let runnerFullName = runnerName+' '+runner[1];  // from col A & B
            return AppendPositionsForResult(runnerFullName,resultRange);
          } else {
            if (debug) Logger.log('No later result needs appended for unique runner, '+runnerNameId);
          }
        } else {
          if (eventDate) Logger.log('No result found for runner, '+runnerName+' on '+eventDate);
          else Logger.log('WARNING: No results found for runner, '+runnerName);
        }
      });
    }))
  })
  .catch(error =>
    Logger.log('ERROR: '+error)
  )
  .finally(() =>
    CloseChromeBrowser()
  );
}

function PasteAllResultsForRunner(
  allResults,
  runnerNameId = 'Sarah_17')
{
  var resultsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(runnerNameId);
  const dateINDEX = 1; // column B
  // Logger.log(typeof allResults)
  // Paste in reverse order from parkrun site => earliest result will be first
  allResults.reverse().forEach(thisResult => {
    thisResult = thisResult.map(CleanValue);  // extracts links into hyperLinks
    thisResult[dateINDEX] = FormatDate(thisResult[dateINDEX],dateFORMAT);
    AppendResultRow(thisResult,resultsSheet);   // applies hyperLinks from Clean
  });
}

/**
 * Imports all results for a given parkrunner ID, pasting them into the sheet.
 * Note: Does NOT append positions (done later in Step 4 catch-up).
 *    @param {string} parkrunnerId - The ID of the parkrunner
 *  @returns {Promise} Resolves when all results are pasted
 * Assumes browser already opened
 */
function ImportAllResultsForRunner(
  parkrunnerId = '11306668',
  runnerNameId = 'Jonah_17',
  thisPage = undefined    // expected known when 'ALL' (since required to get name)
) {
  if (debug) Logger.log('Import All from '+parkrunnerId+'to results sheet, '+runnerNameId);
  CopyResultForRunner(parkrunnerId,'ALL',thisPage)
    .then(allResults => {
      if (allResults) {
        PasteAllResultsForRunner(allResults,runnerNameId);
      }
    })
    .catch(err => 
      Logger.log('ERROR: '+error));
}

/**
 * Creates a new results sheet for a runner if one doesn't exist.
 *    @param {Array of strings} runnerFullName - The full name of the runner
 *    @param {String} ender, Male / Female
 *    @param {String}
 *    @param {String}
 *    @param {number} parkrunnerId - The Park Runner ID
 * @return {String} the name of the new results sheetm based on first name and row number
 */
function CreateRunnerResultsSheet(
  runnerNames = ["Alan","CHAMPION"], gender = 'Male',
  email = '', dob = '19-Oct-1956',  // otherwise null dob if runnerIndex row exists with these details
  parkrunnerId = '777764',
  runnerIndex = undefined   // add  to bottom (index) in Runners sheet unless updating 
)
{
  if (debug) {
    Logger.log('Runner Names: '+runnerNames);
    Logger.log('Gender: '+gender);
    Logger.log('Email: '+email);
    Logger.log('DoB: '+dob);
    Logger.log('Id: '+parkrunnerId);
  }
  if (gender && dob) {    // for new members except the first runner entry 
    rangeRow = allRunnersSHEET.appendRow(
      [...runnerNames,        // A..B
      gender,email,dob,       // C..E
      null,null,null,null,    // F..I derived categories
      parkrunnerId,null,null] // J..L parkrun + derived status: has results. has positions
    );
    runnerIndex = allRunnersSHEET.getLastRow()-runnersStartROW;
  } else {
    // Assume details already on Runners Sheet for first runner
  }
  // Create runner's results sheet if it doesn't exist
  let spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let [runnerName,runnerSurname] = runnerNames;
  let runnerNameId = runnerName+'_'+runnerIndex;
  let templateResults = spreadsheet.getSheetByName(templateNAME);
  let newResultsSheet = templateResults.copyTo(spreadsheet).setName(runnerNameId);
  // ensure the content of the new sheet is unique
    let runnerFullName = runnerNames.join(" ");
    const titleNameCELL = "A1";
    const allRunnersGID = allRunnersSHEET.getSheetId();
    newResultsSheet.getRange(titleNameCELL)
      .setFormula('=HYPERLINK("#gid='+allRunnersGID+'","'+runnerFullName+'")');
  // return to Runners sheet and add the link back and the fast results link
    let newResultsGid = newResultsSheet.getSheetId();
    let parkrunnerResultsUrl = parkrunnerURL+parkrunnerId+allForONE;
    rangeRow = allRunnersSHEET.getLastRow();
    allRunnersSHEET.getRange(rangeRow,1)     // Col. A
      .setFormula('=HYPERLINK("#gid='+newResultsGid+'","'
        +runnerName+'")');
    allRunnersSHEET.getRange(rangeRow,2)     // Col. B     
      .setFormula('=HYPERLINK("'+parkrunnerResultsUrl+'","'
        +runnerSurname+'")');
    if (runnerIndex != 0) {
      let numCols = allRunnersSHEET.getLastColumn();
      allRunnersSHEET.getRange(rangeRow-1,1,1,numCols)
        .copyFormatToRange(allRunnersSHEET,1,numCols,rangeRow,rangeRow);
    }
  return runnerNameId;    // name of the newly created sheet
}

/*
/  Hierarchy for two use-cases of a new runner:
/
/   1.  AddFamilyMember (as a member of your family/club)
/         PromptNewRunner (addCASE)-> with Form to get parkrun id, Dob, email (option)
/       DoAddFamilyMember
/         OpenChromeBrowser->   
/         >GetRunnerResults (to get Name from the Results Page)
/           AccessPage
/         GetRunnerDetails
/         >CreateRunnerResultsSheet (with Name, form details & results from Page)
/         ImportAllResultsForRunner
/           CopyResultForRunner (ALL assumes Results Page preloaded)
/           PasteAllResultsForRunner (chronologically)
/             CleanValue
/             FormatDate
/             AppendResultRow
/         LockCallerForwardsTo-> (triggers)
/           BatchPositionsForRunner...
/
/   2a. SpawnNewFamily (to be owned by first member)
/         PromptNewRunner (spawnCASE)->  with Form to get parkrun id, Dob
/       DoSpawnNewFamily
/         OpenChromeBrowser->
/         >GetRunnerResults (to get Name from the Results Page)
/           AccessPage
/         GetRunnerDetails
/         >InstantiateFamilySpreadSheet (with Family Name)
/         CreateRunnerResultsSheet (with Name, form details but noresults)
/         TriggerFunction => AddFirstMember (with Name & results page)
/
/   2b. AddFirstMember (as a member of new family/club) with Results Page
/         OpenChromeBrowser->
/         >CreateRunnerResultsSheet (with Name)
/         >ImportAllResultsForRunner (assumes details on Runners sheet)
/           CopyResultForRunner (ALL assumes Results Page preloaded)
/           PasteAllResultsForRunner (chronologically)
/             CleanValue
/             FormatDate
/             AppendResultRow
/         LockCallerForwardsTo-> (triggers)
/           BatchPositionsForRunner...
/
*/

function GetRunnerDetails(thisPage) {
  if (!thisPage)
    throw new Error('ERROR: Delayed or unable to access Runner results');
  const h2REGEXP = /<h2>(.*?)<\/h2>/;
  let runnerFullName = (thisPage.match(h2REGEXP) || [])[1]
    .replace(/<span.*?<\/span>/,'')
    .replace(/&nbsp;/g,'')
    .trim();
  if (debug) Logger.log ('3b. Name: '+runnerFullName);
  const paraREGEXP = /<p>(.*?)<\/p>/s;
  let paraContent = (thisPage.match(paraREGEXP) || [])[1];
  // if (debug) Logger.log ('Paragraph: '+paraContent); 
  let category = paraContent.trim().split(' ').pop();
  // if (debug) Logger.log ('Category: '+category);
  var gender = category
    ? ( category.includes('M') ? 'Male'
      : category.includes('W') ? 'Female'
      : null)
    : null;
  return [runnerFullName,gender]; 
}

/**
 * Adds the first Runner to the existing family with basic details
 *  1. Opens the browser session
 *  2. Creates new Results Sheet for first runner 
 *  3. Imports all the results for the new Runner (without Positions)
 *  4. Triggers the Batch process to append positions
 */
function AddFirstMember(
  parkrunnerId,runnerFullName,resultsPage)
{
  if (debug) Logger.log('1. Full name: '+runnerFullName);
  OpenChromeBrowser()
    .then(() => {
      if (debug) Logger.log('2. First runner: '+parkrunnerId);
      const firstINDEX = 0;
      runnerNames = runnerFullName.split(' ');
      return CreateRunnerResultsSheet(
        runnerNames,
        null,
        parkrunnerId,
        null,         // update existing row in Runners sheet or new next row if null
        firstINDEX    // first runner
      );
    })
    .then(runnerNameId => {
      ImportAllResultsForRunner(parkrunnerId,runnerNameId,resultsPage);
      return runnerNameId;    // only this thereafter
    })
    .then((runnerNameId) => {
      let [runnerName,runnerIndex] = runnerNameId.split('_');
      Logger.log('Forking new: '+runnerName+'\t['+runnerIndex+']');
      LockCallerForwardsTo(threadBatchFN,'forked',runnerNameId);
    })
    .catch(err => {
      Logger.log('ERROR: Add First Member, '+parkrunnerId+'\n'+err);
      CloseChromeBrowser();
    })
    .finally(() =>
      // CloseChromeBrowser() // NOT yet until forked process is done!
      Logger.log('Ordinarily, preserve browser session until completed'));
}

function GetTemplateId(templateName) {
  const files = DriveApp.getFilesByName(templateName);
  if (files.hasNext()) {
    return files.next().getId();
  } else {
    throw new Error('Template not found: ' + templateName);
  }
}

const addCASE = {
  title: 'Add Family / Club Member',
  desc:  'This adds a new member parkrunner into the current Spreadsheet and captures their results. '
        +'The runner is identified by a new row in the Runners sheet plus a separate sheet for the'
        +'results of that new parkrun family member (based on their first name and unique row index), '
        +'where the detailed positions of those results will eventually appear from a background task.',
  action: 'Add',
  handler: 'DoAddFamilyMember'
};

function PromptNewRunner(
  thisCase = addCASE)
{
  const ui = SpreadsheetApp.getUi();
  let formHTML = '\n'+
    '<form>\n'+
    '  <div>'+thisCase.desc+'</div><br><br>\n'+
    '  <label>parkrun Id (barcode numeric part only):\t</label>\n'+
    '    <input type="text" id="parkrunnerId"><br><br>\n'+
    '  <label>Date of birth (for accurate %age grade):\t</label>\n'+
    '    <input type="text" id="dob" placeholder="dd-Mmm-YY"><br><br>\n'+
    '  <label>Email (optional, for delegating access):\t</label>\n'+
    '    <input type="text" id="email"><br><br>\n'+
    '  <input type="button" id="submitButton" value="'+thisCase.action+'"\n'+
    '    onclick="document.getElementById(\'submitButton\').disabled=true;\n'+
    '      document.getElementById(\'submitButton\').value=\'Processing...\';\n'+
    '      google.script.run.'+thisCase.handler+'([\n'+
    '      document.getElementById(\'parkrunnerId\').value,\n'+
    '      document.getElementById(\'dob\').value,\n'+
    '      document.getElementById(\'email\').value]);\n'+
    '      google.script.host.close()">\n'+
    '  <input type="button" value="Cancel"\n'+
    '    onclick="google.script.host.close()">\n'+
    '  </form>\n';
  var form = ui.showModalDialog(HtmlService.createHtmlOutput(formHTML), thisCase.title);
  // return [parkrunnerId,dob,email];
}

/**
 * Adds a new Runner to the existing family after prompting for details:
 *  1. Prompts for a new parkrunner id, etc.
 *  2. Opens the browser session
 *  3. Gets the Runners Results Page
 *  4. Creates a new Results Sheet for the new Runner (in same file)
 *  5. Imports all the results for the new Runner (without Positions)
 *  6. Triggers the Batch process to append positions
 */
function AddFamilyMember() {
  PromptNewRunner();    // default addCASE
  // callback to DoAddFamilyMember with [parkrunnerId,dob,email] from form
}

async function DoAddFamilyMember(form) {
  let [parkrunnerId,dob,email] = form;
  if (debug) Logger.log('1. Prompt: '+form);
  let dobRegex = /^\d{2}-[A-Za-z]{3}-\d{2}$/;
  if (!dobRegex.test(dob)) {
    throw new Error('ERROR: Invalid date format - use dd-Mmm-YY');
    // otherwise coerce to this format
  }
  OpenChromeBrowser()
    .then(() => {   // allow time to open browser on server
      if (debug) Logger.log('2. Open: '+parkrunnerId);
      return GetRunnerResultsPage(parkrunnerId);
    })
    .then(resultsPage => {   // after load page in browser
      if (debug) Logger.log('3a. Runner:'+parkrunnerId);
      let [runnerFullName,gender] = GetRunnerDetails(resultsPage);
      let runnerNames = runnerFullName.split(' ');
      if (debug) Logger.log('3b. Details: '+runnerNames+' '+gender+' '+email+' '+dob+' '+' '+parkrunnerId);
      let runnerNameId = CreateRunnerResultsSheet(
        runnerNames,gender,  // to go into cols A & B, C
        email,dob,           // into cols.D & E (hidden for security, as also F..H)
        parkrunnerId);       // into col J (after derived age-category in col. I)
        if (debug) Logger.log('4. Create sheet: '+runnerNameId);
      return [runnerNameId,resultsPage]; // c/fwd resultsPage
    })
    .then(([runnerNameId,resultsPage]) => {  // after create sheet (although same file)
      ImportAllResultsForRunner(parkrunnerId,runnerNameId,resultsPage);
      if (debug) Logger.log('5. Import: '+parkrunnerId+' '+runnerNameId);
      let [runnerName,runnerIndex] = runnerNameId.split('_');
      Logger.log('Added family member with their results: '+runnerName+'\t['+runnerIndex+']');
      LockCallerForwardsTo(threadBatchFN,'added',runnerNameId);
    })
    .catch(err => {
      Logger.log('ERROR: Add Family Member, '+parkrunnerId+'\n'+err);
      CloseChromeBrowser();
    })
    .finally(() =>
      // CloseChromeBrowser() // NOT yet until forked process is done!
      Logger.log('Ordinarily, preserve browser session until completed'));
}

/*
/   2a. SpawnNewFamily (to be owned by first member)
/         OpenChromeBrowser->
/         >PromptNewRunner (Spawn) with Form to get parkrun id, Dob
/           GetRunnerResults (to get Name from the Results Page)
/             AccessPage
/         GetRunnerName
/         >InstantiateFamilySpreadSheet (with Family Name)
/         >TriggerFunction => AddFirstMember (with arguments, Name & Form details)
*/

const templateSHEET = 'FAMILY Template';
const templateFOLDER ='Spawned';
const clubTYPE = 'Parkrunners';   // or 'ClubRunners'

function InstantiateFamilySpreadSheet(
  clubType = clubTYPE,   // or Clubrunners
  runnerNames = ['Peter','WALLIS'],  // into cols A & B of 1st Runners row of new family Spreadsheet
  gender = 'Male',    // for col C
  email = '',         // for Col D
  dob = undefined,    // for Col E (hidden for security, as also F..H)
  parkrunnerId)       // for col J (after derived age-category in col. I)
{
  let targetFolder = DriveApp.getFoldersByName(templateFOLDER).next();
  if (!targetFolder) {
    targetFolder = DriveApp.createFolder(templateFOLDER);
  }
  let templateId = GetTemplateId(templateSHEET);
  if (debug) Logger.log('Template Id: '+templateId);
  let templateFile = DriveApp.getFileById(templateId);
  if (debug) Logger.log('Template File: '+templateFile);
  let familyName = runnerNames[1];
  let familySheetFile = familyName+' '+clubType;
  if (debug) Logger.log('Family sheet: '+familySheetFile);
  let newFamilySpreadsheet = templateFile.makeCopy(familySheetFile,targetFolder);
  let familySheetId = newFamilySpreadsheet.getId();
  if (debug) Logger.log('Family sheet Id: '+familySheetId);
  let ssNew = SpreadsheetApp.openById(familySheetId);
  // pass values here - alternative to pass by argument?
  let runnersSheet = ssNew.getSheetByName(runnersSheetNAME);
  runnersSheet.getRange(1,1).setValue(familySheetFile);   // conveniently file name in A1
  runnersSheet.getRange(runnersStartROW,1,1,5)        // cols A..E ) 5 params
    .setValues([[...runnerNames,gender,email,dob]]);  //  ...ensures MAP is not disturbed
  runnersSheet.getRange(runnersStartROW,parkrunnerIdCOL).setValue(parkrunnerId);
  return [familySheetFile,familySheetId];
}

function TriggerRemote(firstFunction,spawnedSheetId,...args) {
  let spawnedSheet = SpreadsheetApp.openById(spawnedSheetId);
  spawnedSheet[firstFunction](...args);
}

const spawnCASE = {
  title: 'Spawn New Family / Club',
  desc: 'This spawns a new family/club Spreadsheet that captures all results for that first '
        +'member. The surname of this first runner is used to name the new Spreadsheet in which '
        +'they will appear in the first row of the Runners sheet, with a separate sheet '
        +'(&lt;first name&gt;_0) where the results of that first member will eventually appear '
        +'via a background task.',
  action: 'Spawn',
  handler: 'DoSpawnNewFamily'
};

/**
 * Spawns a new Family after prompting for details of the first runner
 *  1. Prompts for a new parkrunner id, etc.
 *  2. Opens the browser session
 *  3. Gets the Runner's Details on their Results Page
 *  4. Instantiates a new Family SpreadSheet
 *  4b.  Adds the first Row on empty Runners Sheet
 *  5. Triggers Add First Family (function) with args in the new instance
 */
function SpawnNewFamily() {
  PromptNewRunner(spawnCASE); 
  // callback to DoSpawnNewFamily with [parkrunnerId,dob,email] from form
}

function DoSpawnNewFamily(
  form = ['21283','30-Jan-69',null]
) {
  let [parkrunnerId,dob,email] = form;
  if (debug) Logger.log('1. Prompt: '+form);
  const dobRegex = /^\d{2}-[A-Za-z]{3}-\d{2}$/;
  if (!dobRegex.test(dob)) {
    throw new Error('ERROR: Invalid date format - use dd-Mmm-YY');
    // otherwise coerce to this format
  }
  OpenChromeBrowser()
    .then(() => {   // allow time to open browser on server
      if (debug) Logger.log('2. Open: '+parkrunnerId);
      return GetRunnerResultsPage(parkrunnerId);
    })
    .then(resultsPage => {   // after load page in browser
      if (debug) Logger.log('3a. Runner: '+parkrunnerId);
      let [runnerFullName,gender] = GetRunnerDetails(resultsPage);
      let runnerNames = runnerFullName.split(' ');
      if (debug) Logger.log('3b. Details: '+runnerNames+' '+gender+' '+email+' '+dob+' '+' '+parkrunnerId);
      let familySheet = InstantiateFamilySpreadSheet(
        clubTYPE,
        runnerNames,gender,  // into cols A & B, C of 1st row of new Runners instance
        email,dob,           // into cols.D & E (hidden for security, as also F..H);
        parkrunnerId);       // into col J (after derived age-category in col. I)
        return [familySheet,runnerFullName,resultsPage];
    })
    .then(([familySheet,runnerFullName,resultsPage]) => {    // after create new Spreadsheet file with new Runners instance
      let [familySheetFile,familySheetId] = familySheet;
      if (debug) Logger.log('4. Instantiate: '+familySheetFile);
      let [familyName,clubType] = familySheetFile.split(' '); 
      // get first name from 'Runners' sheet?
      Logger.log('Spawned new family sheet to add 1st runner in: '+familyName+' ['+clubType+']');
      TriggerRemote('AddFirstMember',familySheetId,parkrunnerId,runnerFullName,resultsPage);
    })
    .catch(err => {
      Logger.log('ERROR: Spawn New Family for parkrunner, '+parkrunnerId+'\n'+err);
      CloseChromeBrowser();
    })
    .finally(() =>
      // CloseChromeBrowser() // NOT yet until forked process is done!
      Logger.log('Ordinarily, preserve browser session until completed'));
}

/**
 *  Returns the first result range with matching positions (typically unknown gender position)
 *    @param {Sheet} resultsSheet - A runner's sheet containing the results (+ header & title)
 *    @param {number} positionCol - The column index (1-based) containing position values
 *    @param {string} positionValue - typically empty cell (value unknown)
 *  @returns {Range} - first results row range that matches (or null if no match)
 */
function FirstMatchRange(resultsSheet,positionCol,positionValue) { 
  let resultsRange = resultsSheet.getDataRange()    // focus on the results table 
    .offset(resultsStartROW-1,0,  // start row 3 = offset 2 (for title and header)
      resultsSheet.getLastRow()-(resultsStartROW-1));   // e.g. 10 rows: 10-(3-1) = 8 results
  var positionValues = resultsRange.offset(0,positionCol-1,resultsRange.getNumRows(),1).getValues();
  var matchIndex = positionValues
    .findIndex(([indexValue]) => indexValue === positionValue);  // e.g. ""
  return (matchIndex !== -1) ? resultsRange.offset(matchIndex,0,1,resultsRange.getNumColumns()) : null;
}

/**
 *  Syncs extra positions for each result for the specified runner.
 *    @param {string} runnerName - typically from the Recurse..., function call
 *  @returns whether more batches to do (aside from the promise itself)
 */
async function SyncPositionsPerRunner(
  runnerNameId = 'Dave_14')    // default unit test (assumes precedent OpenChromeBrowser)
{
  const runnerNameCELL = 'A1';
  const ageCatCELL = 'L3';  //runner may be older since 1st run but gender fixed
  const dateINDEX = 1;      // index for column B
  const genderPosnCOL = 9;  // column I is the Gender position (if present already done)
  const batchSizeMAX = 12;    // estimate batch to catch up on within 5-6 minutes
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let resultsSheet = spreadsheet.getSheetByName(runnerNameId);
  let runnerFullName = resultsSheet.getRange(runnerNameCELL).getValue();
  var resultRange = FirstMatchRange(resultsSheet,genderPosnCOL,"");  // 1st with unknown Gender posn
  if (resultRange) {   //skip this runner if all positions known
    // WARNING: Limit batch of results to catch up (recursively) because 6 mins max per App script!
    let lastResultRow = resultsSheet.getLastRow();
    let promises = [];
    let batchMore = batchSizeMAX;
    while (batchMore) {
      let genderPositionKnown = resultRange.getCell(1,genderPosnCOL).getValue();
      if (!genderPositionKnown) {
        promises.push(AppendPositionsForResult(runnerFullName,resultRange));
        batchMore--;
      }
      if (batchMore) {  // Take care NOT to flow past the last row!!
        if (resultRange.getLastRow() < lastResultRow)
          resultRange = resultRange.offset(1,0);  // continue adding to the batch
        else break; // otherwise, there are no more results beyond the last row!
      }
    }
    if (debug) Logger.log('RunnerId: '+runnerNameId+'\t# of results is '+lastResultRow+' and reached result, '+resultRange.getLastRow());
    let remainToDo = lastResultRow-resultRange.getLastRow();
    var moreBatches = (remainToDo > 0);
    return Promise.all(promises).then(results => {
      let updatesApplied = results.filter(Boolean).length;
      let totalEvents = results.length;
      let morePositions = (moreBatches) ? ' with '+remainToDo+' more positions needed to catch up' : '';
      Logger.log('Runner: '+runnerFullName+'\tUpdates applied: '+updatesApplied
        +' (out of '+totalEvents+' in this batch)'+morePositions);
      if (updatesApplied === 0)
        throw new Error('ERROR: Batching aborted  because no updates');
      return moreBatches;
    });
  } else {
    if (debug) Logger.log('Runner: '+runnerFullName+'\tNo (more) blank positions to catch up');
    return false;   // no more batches
  }
}

const threadBatchFN = recurseBatchFN ='BatchPositionsForRunner';  // threaded & potentially recursed
const lockINDEX = 'lockIndex';
const browserWSEP = 'browserInstance';    // re-use same broser for threaded processes
const runnerNameCOLUMN = "A";     // Runners name in Column A 
const runnerSurnameCOLUMN = "B";  // Runners surname in Column B 
const hasResultsCOLUMN = "K";     // Runner's sheet exists with results (D3:D), and a Parkrunner Id (in col J)
const hasPosnsCOLUMN = "L";       // Runner's sheet (if exists) has Positions up-to-date (I3:I) based on GenderPosn

function CleanupBatch(
  thisScript = threadBatchFN) 
{
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger && trigger.getHandlerFunction() == thisScript) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  PropertiesService.getScriptProperties().deleteProperty(lockINDEX);
  PropertiesService.getScriptProperties().deleteProperty(browserWSEP);
  Utilities.sleep(2000); // Allowing time for triggers & property reset
  CloseChromeBrowser();
  // if (debug)
    Logger.log('Cleared batch status (including triggered scripts) and closed browser');
}

/**
 * Returns the equivalent column by number given a letter - e.g. L means 12
 *    @param {string} letter - column letter (A-Z)
 *  @returns {number} column number
 */
function Column(letter='A') {
  return letter.charCodeAt(0)-64;
}

/**
 * Set this status of the runner in their results row
 *    @param {runnerIndex} indicates the runner to be marked as Positions all done
 *  @returns whether or not all the runners Status is complete 
 */
function AllPositionsDone(allStatus) {
  var allDone = allStatus.every(x=>x);   // since status already a flattened array
  if (debug) Logger.log('Finished: '+allDone);
  return allDone;
}

/**
 * Set this status of the runner in their results row
 *    @param {runnerIndex} indicates the runner to be marked as Positions all done
 *  @returns all of the runners Status is complete 
 */
function MarkRunnerPositionsDone(runnerIndex = 0) {
  var runnerStatusRANGE = allRunnersSHEET.getRange(hasPosnsCOLUMN+runnersStartROW+":"+hasPosnsCOLUMN);
  runnerStatusRANGE.getCell(runnerIndex+1,1).setValue(true);
  let runnersStatus = runnerStatusRANGE.getValues().map(x => x[0]);
  if (debug) {
    let changedStati = runnersStatus.map((status,index) =>
      ({index,status,changed: index === runnerIndex
      })
    );
    Logger.log(JSON.stringify(changedStati,null,2)); // pretty form
  }
  return runnersStatus;
}

/**
 * Get this runner index forwarded before unlocking to allow parallel threads for others
 */
function UnlockCallerForwarded() {
  var lock = LockService.getScriptLock();
  let thisRunnerNameId = PropertiesService.getScriptProperties().getProperty(lockINDEX);
  browserSession = PropertiesService.getScriptProperties().getProperty(browserWSEP);
  // Utilities.sleep(1000); // In case final check (for all runners) might miss this update
  Logger.log('Re-using browser, '+browserSession+' for unique runner, '+thisRunnerNameId);
  lock.releaseLock();   // The lock itself is unique (and unnamed) for this node service
  return thisRunnerNameId;
}

/**
 * Lock any further threaded/recursed for the uniques runner index until received
 */ 
function LockCallerForwardsTo(thisFunction,withReason,thisRunnerNameId) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(12000);  // max time before parallel threads may continue
    PropertiesService.getScriptProperties().setProperty(lockINDEX,thisRunnerNameId);
    PropertiesService.getScriptProperties().setProperty(browserWSEP,browserSession);
    ScriptApp.newTrigger(thisFunction)
      .timeBased()
      .after(1000)
      .create();
  } catch(err) {
    Logger.log('Timeout on '+thisFunction+', '+withReason+' for unique runner Id, '+thisRunnerNameId+':\n'+err);
  }
}

/**
 *  Recursively catch up on missing Positions for each runner from Parkrun site
 *    @params {string} - runnerNameId, passed via locked property setting
 */
function BatchPositionsForRunner(/*runnerNameId*/) {
  const runnerNameId = UnlockCallerForwarded();
// begin
  let [runnerName,runnerIndex] = runnerNameId.split('_');
  runnerIndex = parseInt(runnerIndex);   // ensure a number 
  if (debug)
    Logger.log('Runner: '+runnerName+' ['+runnerIndex+']\tSyncing positions...');
  return SyncPositionsPerRunner(runnerNameId).then((moreToDo) => { 
    // Logger.log('Runner: '+runnerName+' ['+runnerIndex+']\t'+moreToDo+' more to do');
    if (moreToDo) {
      Logger.log('Recursing runner: '+runnerName+'\t['+runnerIndex+']');
      LockCallerForwardsTo(recurseBatchFN,'recursed',runnerNameId);
    } else {
      let runnersStatus = MarkRunnerPositionsDone(runnerIndex);
      if (AllPositionsDone(runnersStatus)) {
        CleanupBatch(threadBatchFN);
        return;
      }
    }  
  })
  .catch(err =>
    Logger.log('ERROR: Failed to trigger recursed batches: '+err)
  )
  .finally(() =>
    // CloseChromeBrowser() // NOT yet until status for every
    Logger.log('Ordinarily, preserve browser session until completed')
  );
}

/**
 * Retrospectively update Positions for each runner via separately triggered processes
 * in parallel, where batching also may be necessary to avoid exceeding 6 seconds max per set.
 */
function CatchUpAllPositions() {
  let runnersStatus = [];
  return OpenChromeBrowser().then(() => {   // Browser always launched beforehand...
    let runners = allRunnersSHEET.getRange(runnerNameCOLUMN+runnersStartROW+":"+runnerNameCOLUMN)
      .getValues().map(x => x[0]).filter(String);
    var runnersResults = allRunnersSHEET.getRange(hasResultsCOLUMN+runnersStartROW+":"+hasResultsCOLUMN)
      .getValues().map(x => x[0]);
    // Ensure ALL threads use the same status so that closure is when done for ALL runners
    runnersStatus = allRunnersSHEET.getRange(hasPosnsCOLUMN+runnersStartROW+":"+hasPosnsCOLUMN)
      .getValues().map(x => x[0]);
    // Thread process for each valid runner in parallel, with a non-conflicting delay 
    runners.forEach((runnerName,runnerIndex) => {
      if (runnersResults[runnerIndex]) {    // if runner has at least one result
        if (!runnersStatus[runnerIndex]) {  // ...and positions not already caught-up
          let runnerNameId = runnerName+'_'+runnerIndex;  // consistently unique for index and status
          Logger.log('Threading runner: '+runnerName+' ['+runnerIndex+']');
          LockCallerForwardsTo(threadBatchFN,'threaded',runnerNameId);
          Utilities.sleep(11000); // delay between activating runner threads in parallel
        }
        else
          Logger.log('Positions up-to-date for runner: '+runnerName+' ['+runnerIndex+']');
      } else {
        if (!runnersStatus[runnerIndex]) {   // Since no results, then already caught-up
          runnersStatus = MarkRunnerPositionsDone(runnerIndex);   // avoids impact on not finishing
        }
      }
    });
  })
  .catch(err => 
    Logger.log('ERROR: Failed to trigger threaded batches: '+err)
  )
  .finally(() => {
    if (AllPositionsDone(runnersStatus)) {
      CleanupBatch(threadBatchFN);
      return;
    }
    else Logger.log('Preserve browser session until all positions updated.')
  });
}

/* ---------------------------------------------------------------------------
/
/   The following definitions and functions are used to support the automatic
/   creation of comparison charts.
/   The SpreadSheet itself has built in functions to support automated charts.
/   These include standard comparative graphs for different groups that may be
/   customised to suit what best inspires a family or club.
/   There are eight samples of filtered charts (mostly by Age-Grade %age):
      Juniors (over past 1.5yrs)
      Seniors (all time)
      Veterans (under 50, over past 2 yrs)
      Supervets (50+,over past year)
      By FAMILY surname (past year)
      Females (by Gender position, over past 2 years)
    Note that these are formulae to filter the selection but can be manually set also  
    By example, here is a sample Group runner selection formula for Veterans under 50

      =LET(
        stage,"Veteran", 
        start,3, 
        ageUnder,50,
        TRANSPOSE(ARRAYFORMULA(
          LET(
            cond,Runners!K:K*(Runners!G:G=stage)*(Runners!F:F<ageUnder),
            {
              FILTER(Runners!A:A,cond), 
              FILTER(ROW(Runners!A:A)-start,cond)
            }
          )
        ))
      )

    On Rankings sheet,there are no function except a common formula,that is the
    same for each ranking table. Note that 0+dateSource ensures a DATE from a
    string (which is the new convention instead of a UTC-compliant date):

      =LET(
        start,3,
        recentYrs,IF(H1="∞",0,H1),
        yearDays,365.25, dateLink,2,
        fTime,5, ageGrade,6, isPB,8,
        genPosn,9, agePosn,10, gradePosn,11, ageCat,12,
        foreNames,UNIQUE(FILTER(Runners!$A$3:$A,Runners!$K$3:$K)),
        sortHeader,J1, headers,B2:J2,
        sortBy,MATCH(sortHeader,headers,0),
        sortOrder,IF(LEFT(sortHeader,1)="↓",FALSE,TRUE),
        SORT(
          BYROW(foreNames, LAMBDA(name,
            LET(
              idx,MATCH(name,Runners!$A:$A,0)-start,
              nameId,name&"_"&idx,
              dateSource, INDIRECT(nameId & "!B:B"),
              timeSource, INDIRECT(nameId & "!E:E"),
              fResult, FILTER( INDIRECT(nameId&"!A:L"),
                IF( recentYrs = 0, timeSource <> "" ,
                  0+dateSource >= TODAY()-recentYrs*yearDays
                )
              ),
              fastTime,MATCH(MIN(
                INDEX(fResult,0,fTime)), 
                INDEX(fResult,0,fTime),0 ),
              {
                name,
                INDEX(fResult,fastTime,dateLink),
                INDEX(fResult,fastTime,fTime),
                INDEX(fResult,fastTime,ageGrade),
                INDEX(fResult,fastTime,isPB),
                INDEX(fResult,fastTime,genPosn),
                INDEX(fResult,fastTime,agePosn),
                INDEX(fResult,fastTime,gradePosn),
                INDEX(fResult,fastTime,ageCat)
              }
            )
          )),
          sortBy,
          sortOrder
        )
      )

    For maintenance needs, here is the hierarchy of functions for the comparative charts:

      GenerateChartsFromGroups
        ExtractGroupRunners
        When referencing a new Performances sheet...
          ClearPerformancesSheet
        GenerateGroupChartInPerformances
          FilterGroupRunnersDatedPerformances
            CollateGroupRunnersDatedPerformances
          CopyGroupPerformancesToSheet
          EmbedGroupPerformancesChart
            ApplyFormatsOnGroupPerformancesCharts

      When any change of hex colour codes in the top Legends table...
        ColourLegendsInGroups
/
/ ---------------------------------------------------------------------------
*/

const datesINDEX = 1;           // for col B on runners' results sheet
const timesINDEX = 4;           // for col E on runners' results sheet
const ageGradesINDEX = 5;       // for col F on runners' results sheet
const chartYAxisTITLE = 'Age Grade';
const ageGradeFORMAT = '0.0%';  // for %age on graphs

/**
 * Returns an array of selective runners performances versus event dates,
 *  ensuring event dates are unique (even if event location differs),
 *  based on most recent years (unless all performances override).
 *    @param {Array<Date>} allDates     - Array of non-unique event dates
 *    @param {Array<string>} runners    - 2D Array of selective runners' names with unique indices
 *    @param {Object} runnersPerfs      - Object of performances for each indexed runner on dates (subset of all dates)
 *    @param {number} [mostRecentYears=recentYRS] - Number of recent years to include (0 to get all)
 *  @returns {Array<Array>} as an Array of arrays containing dated performances (with nulls for absences)
 *           (potentially more efficient as a 2D array)
 */
function CollateGroupRunnersDatedPerformances(allDates,runners,runnersPerfs,
  mostRecentYears = recentYRS)  // assume all performances if zero years cut-off
{
  // Ensure empty performances for runners who missed out
  //    on any run dates (based on recent years only)
  const NOW = new Date();
  let cutoff = mostRecentYears == 0   // simplify the filter (if any)
    ? NOW
    : new Date(
      NOW.getFullYear() - (mostRecentYears|0), 
      NOW.getMonth() - ((mostRecentYears%1)*12|0), 
      NOW.getDate()
    );
  let filteredDates = allDates.filter(date => new Date(date) > cutoff);
  filteredDates.sort((a,b) => new Date(a)-new Date(b));   //sort, oldest first (a-b)
  let uniqueDates = [...new Set(filteredDates)];  // remove duplicates and returnas an Array
  var runnersDatedPerfs = [];
  for (var i=0; i<uniqueDates.length; i++) {
    var row = [uniqueDates[i]];
    for (var j=0; j<runners.length; j++) {
      let runnerNameId = runners[j].join('_');
      if (runnersPerfs[runnerNameId] &&
          runnersPerfs[runnerNameId][uniqueDates[i]] != 0 &&
          runnersPerfs[runnerNameId][uniqueDates[i]] !== undefined) {
        row.push(runnersPerfs[runnerNameId][uniqueDates[i]]);
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
 *    @param {string}                             - chart title,only used for tracking distinct groups
 *    @param {Array<string>}                      - runners - Array of runner names & indices from Group (rarely all)
 *    @param {number} [mostRecentYears=recentYRS] - Number of recent years to include (0 to get all)
 *    @param {number} [perfIndex=ageGradesINDEX]   - Column index of performance values (Age Grade, Time, etc.)
 *  @returns {Array<Array>} Array of arrays containing dated performances (with nulls for absences)
 */
function FilterGroupRunnersDatedPerformances(chartTitle,runners,
  mostRecentYears = recentYRS,   // assume all performances if zero years cut-off
  // performance is normally Age Grade values < 1, presented as %ages,
  //   but potentially Time or Age Grade Posn or # of 1sts may be used
  perfIndex = ageGradesINDEX)
{
  const headerROW = runnersStartROW-1;  // number of header rows above runner's results
  var allDates = [];
  var runnersPerfs = {};
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  runners.forEach(function(runner) { 
    let [runnerName,runnerIndex] = runner;
    if (runnerName.includes("N/A")) {
      Logger.log('WARNING: Missing runners in this group chart ('+chartTitle+'_');
      return;   // for this group chart only
    }
    let runnerNameId = runnerName+'_'+runnerIndex;
    try {
      var resultsSheet = spreadsheet.getSheetByName(runnerNameId);
      if (!resultsSheet) {
        SpreadsheetApp.getUi().alert('Error',
          'No results in unique runner sheet, '+runnerNameId,
          SpreadsheetApp.getUi().ButtonSet.OK);
        return;
      }
      var resultsRange = resultsSheet.getDataRange()
        .offset(headerROW,0,resultsSheet.getLastRow()-headerROW);
      var dates = resultsRange.getDisplayValues()
        .map(date => date[datesINDEX]); // Dates in col B only
      var perfs = [];
      if (perfIndex == timesINDEX) {     //  DisplayValues may be a problem on chart min/max
        perfs = resultsRange.getDisplayValues()
          .map(time => time[timesINDEX]); // Duration times in col E only
      } else {
        perfs = resultsRange.getValues()
        .map(perf => perf[perfIndex]); // Positions or %ages (<1) in col. D, F, or I..K
      }
      runnersPerfs[runnerNameId] = {};
      for (var i=0; i<dates.length; i++) {
        runnersPerfs[runnerNameId][dates[i]] = perfs[i];
      }
      allDates = allDates.concat(dates);
    } catch (err) {
      Logger.log("ERROR: filtering results for unique runner, " +runnerName+'['+runnerIndex+"]\n"+err);
    }
  });
  runnersDatedPerfs = CollateGroupRunnersDatedPerformances(
    allDates,runners,runnersPerfs,mostRecentYears);
  Logger.log('Collate Runners Dated Performances: '+runners);
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
 *    @param {string} [dateFormat=dateFORMAT]   - Format used for displayed dates
 *    @param {number} [decimalPlaces=4]         - Floating precision (unless integer)
 * @returns {Range} where performances are on the sheet (beyond subsequent charts)
 */
function CopyGroupPerformancesToSheet(
  perfsSheet,runners,runnersPerfs,
  groupPerfsRow = 2,
  groupPerfsCol = 4,
  dateFormat = dateFORMAT,     // aligned  with display format
  decimalPlaces = 4)
{
  var dates = ['Date', ...runnersPerfs.map(date => date[0])];
  var groupTable = [dates];   // 'Date' with d-MMM-YY dates are headers
  for (var i=1; i<=runners.length; i++) {
    var runnerPerfs = runnersPerfs.map(result => result[i]);
    var [runnerName,runnerIndex] = runners[i-1];
    var runnerNameId = runnerName+'_'+runnerIndex;  // strip Index later if required?
    var perfRow = [runnerNameId];
    runnerPerfs.forEach(value => {
      var formattedValue;
      if (value === undefined || value === null)
        formattedValue = null;
      else {
        formattedValue = Number(value);
        if (isNaN(Number(value)))  // retrieved as a display value => assume time
          formattedValue = `0:${value}.000`;
        else if (formattedValue % 1 !== 0)
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
  groupPerfsRange.offset(0,1,1,
    groupPerfsRange.getLastColumn()-groupPerfsRange.getColumn()-1)
    .setNumberFormat(dateFormat);
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
function ApplyFormatsOnGroupPerformancesChart (
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
 *    @param {string} chartTitle    - The title of the chart
 *    @param {range} groupPerfsRange      - Performances range on RHS of chart position
 *    @param {Array<Array<string>>} runnersLegend - A 2D array of names, indices & colours.
 *    @param {number} [perfChartRow=2]    - The row number to place the chart at
 *                                           (avoid subsequent charts coinciding)
 *    @param {number} [perfChartCol=2]    - The column number to place the chart 
 *                                           (typically starting at B2)
 *    @param {boolean} [showDates=false]  - Allow non-contiguous dates on the horizontal
 *    @param {boolean} [reverseTrend=1]   - Trends in reverse: rising to low values
 *    @param {boolean} [stripIndex=false] - Simplify legend if first name context unique
 *    @param {number} [filterRecentYears=recentYRS] - The cut-off years for results 
 *    @param {string} [perfTitle=chartYAxisTITLE]   - Performance title on vertical
 *                                                    (typically Age Grade in results)
 *    @param {string} [perfFormat=ageGradeFORMAT]   - Format the performance (e.g. %age)
 */
function EmbedGroupPerformancesChart(
  perfSheet,chartTitle,groupPerfsRange,runnersLegend,
  perfChartRow = 2,
  perfChartCol = 2,
  showDates=false,reverseTrend=1,stripIndex=false,
  filterRecentYears = recentYRS,
  perfTitle = chartYAxisTITLE,
  perfFormat = ageGradeFORMAT)
{
  const chartWIDTH = 800;
  const chartHEIGHT = 350;
  const offsetBORDER = 5; // pixels
  var perfLimits = ApplyFormatsOnGroupPerformancesChart(
    perfSheet,groupPerfsRange,runnersLegend,
    perfChartRow,perfChartCol,perfFormat,
    chartHEIGHT,chartWIDTH,offsetBORDER);
  if (filterRecentYears > 0)
    chartTitle += ' (max '+filterRecentYears+' years)';
  perfSheet.getRange(perfChartRow-1,perfChartCol).setValue(chartTitle);
  if (stripIndex) {  // tweak to remove the unique index if desired on chart legend
    let numRunners = groupPerfsRange.getValues().slice(1).length;
    groupPerfsRange.offset(1,0,numRunners,1)
      .setValues(groupPerfsRange
        .getValues()
        .slice(1)
        .map(name => [name[0].split('_')[0]])
      );
  }
  let colours = runnersLegend.map(colour=>colour[2]); // 3rd col
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
    .setOption('title',chartTitle)
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
function ClearPerformancesCharts(
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
 * Extracts runner names and corresponding indices and colours from a range.
 *    @param {Range} runnersRange - contains runners' names, indices & colours in 3 rows
 * @returns {Array<Array<string>>} 2D array of [name,index,color] for each runner.
 */
function ExtractGroupRunners(runnersRange) {
  var groupRunners = runnersRange.getValues();
  var runnersNames = groupRunners[0];
  var runnersIndices = groupRunners[1];
  var runnersColours =  groupRunners[2];
  var runnersLegend = [];
  for (var j=0; j<runnersNames.length; j++) {
    if (runnersNames[j] != "") {
      runnersLegend.push([runnersNames[j],runnersIndices[j],runnersColours[j]]);
    }
  }
  return runnersLegend;   // 2-D array
}

/**
 * Generate a performance chart in a specified sheet for a group of runners
 *    @param {string} perfSheet     - The name of the sheet to generate the chart in
 *    @param {string} chartTitle    - The full title of the chart
 *    @param {Array<Array<string>>} runnersLegend   - A 2D array runner,index,colour triple
 *    @param {number} perfChartRow  - The row number to position the chart
 *    @param {number} perfChartCol  - The column number to position the chart.
 *    @param {string} showDates     - Allows for long gaps (false unless true)
 *    @param {string} reverseTrend  - Reverse vertical axis (1, unless -1)
 *    @param {string} stripIndex    - Reverse vertical axis (1, unless -1)
 *    @param {number} filterRecentYears - Cut-off excess years (0 if no filter)
 *    @param {string} perfColumnTitle   - The vertical title (default: Age Grade)
 *    @param {number} perfColumnIndex   - The index of the performance  (5 = F)
 *    @param {string} perfFormat        - The format to apply to performances
 */
function GenerateGroupChartInPerformances(
  perfSheet,chartTitle,runnersLegend,
  perfChartRow,perfChartCol,
  showDates=false,reverseTrend=1,stripIndex=false,
  filterRecentYears,
  perfColumnTitle=chartYAxisTITLE,perfColumnIndex,perfFormat)
{
  if (!perfSheet) return null;
  let runners = runnersLegend.map(runner=>[runner[0],runner[1]]); // include unique Id
  // runners are a subset! so index is EITHER from allRunners sheet OR in Legend
  var runnersDatedPerfs = FilterGroupRunnersDatedPerformances(
    chartTitle,runners,filterRecentYears,perfColumnIndex);
  if (!runnersDatedPerfs) return null;
  var groupPerfsRange = CopyGroupPerformancesToSheet(perfSheet,runners,
    runnersDatedPerfs,perfChartRow);
  if (!groupPerfsRange) return null;
  var perfChart = EmbedGroupPerformancesChart(
    perfSheet,chartTitle,groupPerfsRange,runnersLegend,
    perfChartRow,perfChartCol,
    showDates,reverseTrend,stripIndex,
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
function GenerateChartsFromGroups(
  groupsSheetName = "Groups")
{
  var groupsSheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName(groupsSheetName);
  if (!groupsSheet)
return;
  const numGroupROWS = 4;       // Group of related runners info spread over 4 rows 
  const numRunnerROWS = numGroupROWS-1;
  const groupsStartROW = 6;     // Title & header rows sandwich 3-row loopup table
  const runnersStartCOL = 3;     // Output sheet name & Group titleprecede  
  const groupsEndROW = groupsSheet.getLastRow();
  const maxGroupsCOUNT = parseInt((groupsEndROW-groupsStartROW+1)/numGroupROWS);
  const maxRunnersCOUNT = groupsSheet.getLastColumn()-runnersStartCOL+1;
  const paramsCOUNT = 10;       // Assume parameters of Group in Columns A..J
  var defaultPerfSheet = null;  // Assume default Performance Sheet unless specified
  var is1stGroup = true;
  // For each Group, there are THREE ordered steps to generate the chart...
  for (var i=0; i<maxGroupsCOUNT; i++) {
    var group1stRow = groupsStartROW+numGroupROWS*i; // Assumes no blank rows between
    
    // 1. Establish valid Group in runners (2D-Array) with colours on two rows...
    var runnersRange = groupsSheet.getRange(group1stRow+1,  // Runner names on 2nd row
      runnersStartCOL,numRunnerROWS,maxRunnersCOUNT);   // incl. runner index & legend (2..4)
    var runnersLegend = ExtractGroupRunners(runnersRange);
    if (!runnersLegend || runnersLegend.length == 0) continue; // skip after step 2 if no runners

    // 2. Extract Group parameters from 1st row, with chart position in Col. A of 2nd row
    //    Assume null for function defaults (if empty) - see notes on Groups sheet
    var paramsRange = groupsSheet.getRange(group1stRow,1,1,paramsCOUNT);
    var perfSheet,perfSheetName,groupName,groupTitle,
      showDates,reverseTrend,stripIndex,
      filterRecentYears,
      perfColumnTitle,perfColumnIndex,perfFormat;
    var params = paramsRange.getValues()[0];
    for (var j=0; j<params.length; j++) {   // ensure all parameters considered
      switch (j) {
      case 0:     // column A (of Groups sheet)
        perfSheetName = params[0] || defaultPerfSheet;
        if (is1stGroup || perfSheetName != defaultPerfSheet) {  // target is 'Performance' sheet
          perfSheet = ClearPerformancesCharts(perfSheetName);
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
        // used for vertical & main titles , AND
        // ... criteria matches a header in results sheets
        perfColumnTitle = params[3] || chartYAxisTITLE;  break;
      case 4:     // column E
        perfColumnIndex = params[4] || undefined;     break;
      case 5:     // column F
        perfFormat = params[5] || undefined;          break;
      case 6:     // column G
        filterRecentYears = params[6] || undefined;   break;
      case 7:     // column H
        showDates = params[7] == true || false;       break;
      case 8:     // column I - TODO ineffective on reverse
        reverseTrend = params[8] == true ? -1 : 1;    break;
      case 9:     // column J 
        stripIndex = params[9] == true || false;      break;
      }
    } // end of parameters for one Group from Groups sheet
    var perfChartRow = paramsRange.offset(1,0)    // in col A on 2nd row
      .getValue() || undefined;  
    var perfChartCol = paramsRange.offset(1,1)    // in col B (default col B)
      .getValue() || undefined;
    var firstRunner = paramsRange.offset(1,2) 	  // in col C
      .getValue() || undefined;  
    if (!groupName || firstRunner === "#N/A")
      continue;   // skip if no runners (after having cleared since last run)

    // 3. Place Group performances data on sheet before creating Group chart
    var chartTitle = groupName+" ("+groupTitle+") "+perfColumnTitle;
    GenerateGroupChartInPerformances(
      perfSheet,chartTitle,runnersLegend,
      perfChartRow,perfChartCol,
      showDates,reverseTrend,stripIndex,
      filterRecentYears,
      perfColumnTitle,perfColumnIndex,perfFormat);
  } // end of Groups from Groups sheet
  return perfSheet;
}

function ColourLegendsInGroups(
  sheetName="Groups")
{
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
          sheet.getRange(i+1,j+1).setFontColor("#ffffff");  // TODO or #000000 for contrast?
        }
      }
    }
  }
}

function GetRelatedTabColor(
  nameIndex = 'Alan_13')
{
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  let resultsSheet = spreadSheet.getSheetByName(nameIndex);
  const colour = resultsSheet.getTabColor();
  Logger.log('Colour: '+colour);
  return colour || "#ffffff"; // default if no color
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
  var parkrunsMenu = ui.createMenu('parkrun')
    .addItem("Import result for each runner"+
      "\u00A0".repeat(16)+"Ctrl+Alt+Shift+0",
      'ImportResultForEachRunner')
    .addItem("Generate charts from Groups"+
      "\u00A0".repeat(15)+"Ctrl+Alt+Shift+1",
      'GenerateChartsFromGroups')
    .addSeparator()
    .addItem("Protect results sheets per runner"+
      "\u00A0".repeat(9)+"Ctrl+Alt+Shift+3",
      'ReprotectEachRunnerResultsSheets')
    .addItem("Colour legends in Groups"+
      "\u00A0".repeat(22)+"Ctrl+Alt+Shift+4",
      'ColourLegendsInGroups')
    .addItem("Catch-up all positions"+
      "\u00A0".repeat(28)+"Ctrl+Alt+Shift+6",
      'CatchUpAllPositions')
    .addSeparator()
    .addItem("Add family (or club) member"+
      "\u00A0".repeat(16)+"Ctrl+Alt+Shift+7",
      'AddFamilyMember')
    .addItem("Spawn new family (or club)"+
      "\u00A0".repeat(19)+"Ctrl+Alt+Shift+9",
      'SpawnNewFamily')
    // .insertMenu(ui,5)   // ideally before Tools 
    .addToUi();
}

function PasteAboveRangeFormula() {
  var sheet = SpreadsheetApp.getActiveSpread().getActiveSheet();
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
