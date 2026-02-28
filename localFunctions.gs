/* -------------------------------------------------------------------------------------
/
/ This is a client-end GoogleApp Script that resides within the Family Template
/ Google Spreadsheet.  The scope here is limited to local synchronous  functions 
/ that are NOT dependent on GCR functions.  These wer developed as part of phase I
/ although significantly upgraded as a consequence of later phases .
/ The primary five entry-point functions that are bound to macros/keys are:
/   1.  Protect results for each runner
/       (ProtectEachRunnerResultsSheets - Ctrl+Alt+Shift+2)
/   2.  Generate charts from Groups
/       (GenerateChartsFromGroups       - Ctrl+Alt+Shift+1)
/ 	3.  Colour legends in Groups
/       (ColourLegendsInGroups)
/   4.  Clean format for results)       - => import in GCR functions?
/       (CleanFormatforPastedRunResults - Ctrl+Alt+Shift+0)
/   5.  Scroll to last result
/       (ScrollToLastResult             - Press down arrow in M3)
/---------------------------------------------------------------------------------------
 */

/**
 * @OnlyCurrentDoc
 *  // Ensure authorisation granted via appsscript.json
 * @scope https://www.googleapis.com/auth/script.external_request
 * @scope https://www.googleapis.com/auth/script.scriptapp
 * @scope https://www.googleapis.com/auth/spreadsheets
 * @scope https://www.googleapis.com/auth/script.container.ui
 */

const resultTABLE = "Event"   // For any Runner Event results sheet
const eventHeaderCELL = "A2"; //  where the header row is below the Runner name title
const firstResultCELL = "A3"; //  and at least one Result has been cleanly entered
const firstResultROW = 3;     //  where new Results MUST be below this first result
const scrollCOL = 13;      // Scroll down when selecting column M beyond the table
const pasteCOL = 1;        // Paste Event result starting in 1st column (A)
const timeCOL = 5;         // Event result time is in the 5th column (E)

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
/       EnsureBlankResultsRange (range for new results)
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
  const addersCOLUMN = "D";     // Where the results adder(s) are in the Runners table
  var indexRow = allRunnersSheet.getRange("A3:A").getValues().map(function(value) { 
    return value[0];
  }).indexOf(runnerName);
  if (indexRow == -1) {
    Logger.log("Individual runner, "+runnerName+" not found within the Runners sheet table");
    return [];
  }
  var addersCell = allRunnersSheet.getRange(addersCOLUMN+(indexRow+runnersStartROW));
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
    Logger.log("ERROR: adding editor: "+editor+". ERROR: "+err+" because not a google email address");
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
    SpreadsheetApp.getUi().alert('WARNING',"Without protection, any spreadsheet Editor can add results to "+
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
  var runnerFullNames = allRunnersSheet.getRange(runnersRANGE)
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
    var resultsSheet = activeSpreadsheet.getSheetByName(runnerNameId);
    if (!resultsSheet) {
      var thisRow = index+runnersStartROW;
      var parkrunnerId = allRunnersSheet.getRange(
        thisRow,parkrunnerIdCOL).getValue();
      // resultsSheet = CreateRunnerResultsSheet(runnerFullName,parkrunnerId);
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
  var sheet = activeSpreadsheet.getActiveSheet();
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
    same for each ranking 2D table. Note that 0+dateSource ensures a DATE from a
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
              fResult,FILTER(INDIRECT(nameId&"!A:L"),
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

    Likewise, here are the three formulae for key columns (for the 1st challenge):

    "Prev Best" (determines the target time, based on past n years, where n is in B$1)
      =LET(
        start,3,
        recentYRS,B$1,
        MAP($A3:$A,LAMBDA(name,
          LET(
            name_Id,name&"_"&ROW(name)-start,
            IFERROR(MIN(
              FILTER(
                INDIRECT(name_Id&"!E:E"),
                (0+INDIRECT(name_Id&"!B:B")
                  >= TODAY()-recentYRS*365.25)*
                (0+INDIRECT(name_Id&"!B:B")
                  < E$1),
                ISNUMBER(INDIRECT(name_Id&"!E:E"))
              )
            ),"")
          )
        ))
      )

    "Time" (if participated)
      =LET(
        start,3, 
        MAP($A3:$A,LAMBDA(name,
          LET(
            name_Id,name&"_"&ROW(name)-start,
            IFERROR(
              INDEX(
                INDIRECT(name_Id&"!E:E"),
                MATCH(
                  E$1,
                  0+INDIRECT(name_Id&"!B:B")
                ,0)
              ),
              "DNS"
            )
          )
        ))
      )

    "Age Grade %age" (if participated)
      =LET(
        start,3, 
        MAP($A3:$A,LAMBDA(name,
          LET(
            name_Id,name&"_"&ROW(name)-start,
            IFERROR(
              INDEX(
                INDIRECT(name_Id&"!F:F"),
                MATCH(
                  E$1,
                  VALUE(INDIRECT(name_Id&"!B:B")),
                  0
                )
              ),
              "DNS"
            )
          )
        ))
      )

    Thereafter, each challenge block may be hiddenm when a new block is added with a new date.

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
  runners.forEach(function(runner) { 
    let [runnerName,runnerIndex] = runner;
    if (runnerName.includes("N/A")) {
      Logger.log('WARNING: Missing runners in this group chart ('+chartTitle+'_');
      return;   // for this group chart only
    }
    let runnerNameId = runnerName+'_'+runnerIndex;
    try {
      var resultsSheet = activeSpreadsheet.getSheetByName(runnerNameId);
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
  var perfSheet = activeSpreadsheet.getSheetByName(perfSheetName);
  if (!perfSheet) {
    perfSheet = activeSpreadsheet.insertSheet(perfSheetName);
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
  var groupsSheet = activeSpreadsheet.getSheetByName(groupsSheetName);
  if (!groupsSheet) return;
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
          perfSheet = activeSpreadsheet.getSheetByName(perfSheetName);
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
  var sheet = activeSpreadsheet.getSheetByName(sheetName);
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
  let resultsSheet = activeSpreadsheet.getSheetByName(nameIndex);
  const colour = resultsSheet.getTabColor();
  Logger.log('Colour: '+colour);
  return colour || "#ffffff"; // default if no color
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
  var selection = activeSpreadsheet.getActiveRangeList().getRanges();
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
  // activeSpreadsheet.getActiveRangeList().removeAllRanges();
}
