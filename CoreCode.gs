/* -------------------------------------------------------------------------------------
/
/ This is a client-end Google App Script that resides within the Family Template
/ Google Spreadsheet.  The scope here is limited to local synchronous functions 
/ that are NOT dependent on GCR functions.  These werw developed as part of phase I
/ although significantly upgraded as a consequence of later phases .
/ The primary five entry-point functions that are bound to macros & keys are:
/   1.  Protect results for each runner (simplified to latest results only)
/       (ProtectEachRunnerResultsSheets - Ctrl+Alt+Shift+4)
/ 	2.  Colour legends in Groups (distinct for all males and females)
/       (ColourLegendsInGroups          - Ctrl+Alt+Shift+5)
/   3.  Generate charts from Groups (executed in 4 batches)
/       (GenBatchChartsFromGroups       - Ctrl+Alt+Shift+6)
/   4.  Append non-parkrun result
/       (AppendNonParkrunResult)
/   5.  Prepare App (for prototype and Web App, triggered after 3. Generate)
/       (PrepareAppSheets)
/---------------------------------------------------------------------------------------
 */

/**
 * @OnlyCurrentDoc
 *  // Ensure authorisation granted via appsscript.json
 * @scope https://www.googleapis.com/auth/script.external_request
 * @scope https://www.googleapis.com/auth/script.scriptapp
 * @scope https://www.googleapis.com/auth/spreadsheets
 * @scope https://www.googleapis.com/auth/script.container.ui
 * @scope https://www.googleapis.com/auth/userinfo.email
 * 
 */

// Global constants and varaibles must be defined within a this file IF potentially a triggered
// TODO: Sync gcrFunctions.gs with localFunctions.gs
const lc = {
  debug: false,            // WARNING: may slow down performance
  resultsStartROW: 3,
  locationINDEX: 0,       // column A on each Results sheet
  dateINDEX: 1,           // column B on each Results sheet
  timeINDEX: 4,           // column E on each Results sheet
  ageGradeINDEX: 5,       // column F on each Results sheet
  eventCOL: 1,            // column A is the event location
  dateCOL: 2,             // column B is the date of event
  runNumCOL: 3,           // column C is the instance # at event
  groupPerfsCOL: 4,       // column D on each Performances chart sheet (e.g. Leagues)
  PBtickBoxCOL: 8,        // column H tick-box on Results sheets, ticked if value in G is PB
  PBtoAgeCatNumCOLS: 5,   // cols H..L for I..K positions between derived cols, H & L
  genderPosnCOL: 9,       // column I on each Results sheet (until on parkrun results page?)
  ageCatPosnCOL: 10,      // column J on each Results sheet
  ageGradePosnCOL: 11,    // column K on each Results sheet
  ageCatCOL: 12,          // column L is the Age Category on event date (derived from DoB)
  pasteCOL: 1,            // Paste Event result starting in 1st column (A)
  dateFORMAT: 'd-MMM-yy',
  parkrunDateFORMAT: 'dd/MM/yyyy',  // perhaps for UK-based runners?
  universalDateFORMAT: 'yyyy-MM-dd',
  ageGradeFORMAT: '0.0%', // for %age on graphs
  chartYAxisTITLE: 'Age Grade',
  numGroupROWS: 4,        // Group of selected runners info in 4 rows 
  numRunnerROWS: 3,       // ...within group
  groupsStartROW: 9,      // Title, header + lookup legends by gender 
  runnersStartCOL: 3,     // Performance Output sheet start column
  runnersStartROW: 3,
  runnersSheetNAME: 'Runners',
  runnersNameCOLUMN: "A",
  resultTABLE: 'Event',   // First Header title in Results sheet
  eventHeaderCELL: "A2",  // Header row is below the Runner name title
  firstResultCELL: "A3",  //  ...and at least one Result has been cleanly entered
  offsetBORDER: 5,        // number pixels around chart
  groupStringsCOUNT: 5,   // String parameters from Groups sheet up to and including Compare Format
  groupValuesCOUNT: 6,    // ...with remaining parameters are numbers or boolean flags
  paramsCOUNT: 11,        // main header in Groups sheet (runner parameters start in col 3 below)
  recentYRS: 2,           // filter comparison graphs based to most recent years only
  chartTimeSECS: 90,      // time needed to generate a chart (<10 runners?)
  maxTimeSECS: 6*60,      // limit off 6 minutes per GAS script (governs need to batch)
  chartColOFFSET: -2,     // column B from D on any performances chart sheet (e.g. for Leagues)
  trendColOFFSET: 12      // column N from B on runner's own trend chart sheet
};
var lv = {
  activeSpreadsheet: SpreadsheetApp
    .getActiveSpreadsheet(),      // allows dynamic shift to new context
  startTime: Date.now()
};
lv.activeSpreadsheetId = lv.activeSpreadsheet.getId();       // the originator ID
lv.activeSpreadsheetName = lv.activeSpreadsheet.getName();
lv.allRunnersSheet = lv.activeSpreadsheet
  .getSheetByName(lc.runnersSheetNAME);     // ...for new runners from initiating Spreadsheet
if (lc.debug)
  Logger.log('Current Spreadsheet Id: '+lv.activeSpreadsheetId);  
const scrollCOL = 13;         // Scroll down when selecting column M beyond the table

/* ---------------------------------------------------------------------------
/
/   The following definitions and functions support the automatic protection
/   of each runner's result sheets to allow safe entry of results by others.
/   There are two use-case heirarchies of functions:
/
/     1.  ReprotectResultsSheetPerRunner
/           ReprotectResultsSheet*
/             UnprotectResultsSheet (legacy if any)
/             GetProtectResultsRanges (defined on Runners sheet)
/             ReprotectResultsRanges <------------¬
/               UnprotectResultsRangesOnSheet
//              After freeze-protecting earlier results...
/                 GetResultsEditorsForRunner
/                 ProtectResultsRangeByEditor (if any prior)
/
/     2a. AppendNonParkrunResult
/           GetResultsEditorsForRunner
//          If Editor permitted...
/             AppendNewResultRow (for manual entry)
/
/     2b. onEditDetectNeedToReprotect
/           After result entry has completed...
/             ReprotectResultsSheet
/               (*as in 1. above)----------^
/
/  ---------------------------------------------------------------------------
*/

/**
 * Gets two range of results rows to be protected on the active sheet.
 *    @param {number} numRowsChange - max number of rows changeable by Editors
 *  @return {Array of strings} - two distinct ranges to protect: e.g. A1:W,X1:Z  
 */
function GetProtectResultsRanges(
  numRowsChange = maxChangeROWS)
{
  const resultsSheet = SpreadsheetApp.getActiveSheet();
  const runnerNameId = resultsSheet.getName();
  let lastColumn = resultsSheet.getLastColumn();
  let lastRow = resultsSheet.getLastRow();
  let maxRowsChange = Math.min(numRowsChange,lastRow-lc.resultsStartROW);  // dividing line...
  let lastRowFreeze = lastRow-maxRowsChange;
  let rangeFreeze = resultsSheet
    .getRange(1,1,lastRowFreeze,lastColumn)                             // ...all above
    .getA1Notation();                       
  let firstRowChange = lastRowFreeze+1;
  let rangeChange = (maxRowsChange > 0)
    ? resultsSheet
        .getRange(firstRowChange,1,maxRowsChange,lastColumn)            // ...all below (if any)
        .getA1Notation()
    : undefined;
  Logger.log("Protecting edit ranges on results sheet, "+runnerNameId+"...\n"+
    'Owner-only range: '+rangeFreeze+'\n'+
    'Editors range: '+rangeChange+' ('+maxRowsChange+' latest results)');
  return [rangeFreeze,rangeChange];
}

/** LEGACY / SAFEGUARD
 * Unprotect entire sheet if any protection (legacy solution interferes)
 */
function UnprotectResultsSheet(resultsSheet) {
  var protections = resultsSheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => p.remove());
}

/** OBSOLETE
 * Protect entire sheet with exception (legacy solution OBSOLETE)
 */
function ProtectResultsSheetWithException(resultsSheet,exceptionRange) {
  var sheetProtection = resultsSheet.protect();
  // Only the worksheet owner is permitted to allow this range exception 
  var protectionRange = resultsSheet.getRange(exceptionRange);
  // Exception range to be restricted thereafter
  sheetProtection.setUnprotectedRanges([protectionRange]);
}

/**  OBSOLETE
 * Reallows editing on a range of a protected sheet  (legacy solution OBSOLETE)
 *  Removes existing protection, adds an exception for the specified range,
 *  and reapplies protection with domain edit disabled.
 *    @param {string} rangeNotation - The A1 notation of the range to allow editing.
 */
function ReallowRunnerResultsSheetWithException(rangeNotation) {
  let resultsSheet = SpreadsheetApp.getActiveSheet();
  UnprotectResultsSheet(resultsSheet);
  ProtectResultsSheetWithException(resultsSheet,rangeNotation);
  let runnerNameId = resultsSheet.getName();
  if (lc.debug)
    Logger.log("Reprotected results in sheet, "+runnerNameId+
      " sheet, except for the new results range, "+rangeNotation);
}

/**
 * Gets the list of users permitted to update results for a specific runner.
 *    @param {string} runnerNameId - The unique indexed name of the runner (e.g. Tobias_0)
 *  @return {string[]} An array of users (by email) permitted to update latest results.
 */
function GetResultsEditorsForRunner(runnerNameId) {
  const editorsCOL = 4;     // col D in Runners sheet has permitted results editor(s)
  const [runner,index] = runnerNameId.split('_');
  if (isNaN(index)) {
    Logger.log('ERROR: '+runnerNameId+' is not a valid results sheet');
    return null;
  } else {
    var editorsCell = lv.allRunnersSheet.getRange(runnersStartROW+parseInt(index),editorsCOL);
    if (lc.debug)
      Logger.log(editorsCell.getA1Notation());    // e.g. "D16" for runner_index, Alan_13
    var editorEmails = editorsCell.getValue().split(",");
    editorEmails = editorEmails.map(email => email.trim());
    if (lc.debug)
      Logger.log("Latest results on sheet, "+runnerNameId+" may be added/changed by:\n"
        +editorEmails.join(","));
    return editorEmails;
  }
}

/**
 * Remove any range protections on a runner's result sheet (typically two)
 */
function UnprotectResultsRangesOnSheet(resultsSheet) {
  var protections = resultsSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  protections.forEach(function(protection) {
    protection.remove();
  });
}

/**
 * Allow the new range protection on a runner's sheet
 */
function ProtectResultsRangeByEditor(editor,rangeProtection) {
  try {
    rangeProtection.addEditor(editor);
  } catch (err) {
    Logger.log("ERROR: Adding editor, "+editor+"; a google (group) email is expected\n"+err);
  }
}

/**
 * Reprotect a results sheet, allowing a runner to add/change latest results only.
 *    @param {string} rangeFreeze - The range that is to be frozen, e.g. A1:W
 *    @param {string} rangeChange - The range that may be updated, e.g. X1:Z
 *  @return {string[]} An array of runners (emails) permitted to apply changes.
 */
function ReprotectResultsRanges(rangeFreeze,rangeChange) {
  const resultsSheet = SpreadsheetApp.getActiveSheet();
  UnprotectResultsRangesOnSheet(resultsSheet);
  let freezeRange = resultsSheet.getRange(rangeFreeze);
  let thisProtection = freezeRange
    .protect()          // consider: freezeRange.setWarningOnly(true)
    .setDescription('Freeze past results');
  if (rangeChange) {    //  when only one result, prevent any edits
    let changeRange = resultsSheet.getRange(rangeChange);
    thisProtection = changeRange
      .protect()        // default: edit rights limited to Owner
      .setDescription('Edit latest result(s)');
    let runnerNameId = resultsSheet.getName();
    let editorsEmails = GetResultsEditorsForRunner(runnerNameId);
    if (editorsEmails && editorsEmails.length > 0) {
      editorsEmails.forEach(email => {        // edit rights extended to specific Editor
        ProtectResultsRangeByEditor(email,thisProtection);
      });
      if (lc.debug)
        Logger.log("Changes permitted in range, "+changeRange+" only, on results sheet, "+
          runnerNameId+", allowing changes by:\n"+(editorsEmails.join(",")));
    } else {
      SpreadsheetApp.getUi().alert(
        'WARNING',"Without protection, any Editor can change any results in sheet, "+
        runnerNameId,SpreadsheetApp.ButtonSet.OK);
      Logger.log("WARNING: The results sheet, "+runnerNameId+" is unprotected");
    }
    return editorsEmails;
  } else
    return undefined;
}

/** 
 *  For the named runner's result sheet...
 *    0.  Get that runner's results sheet
 *    1.  Unprotect results sheet (compatible with former solution)
 *    2.  Determine owner/editor protection on ranges that may apply 
 *    3.  Apply protection on results, either frozen (by owner) or changed by editors
 */
function ReprotectResultsSheet(
  runnerNameId = 'Alan_13')
{
  const resultsSheet = activeSpreadsheet.getSheetByName(runnerNameId);
  if (resultsSheet) {
    resultsSheet.activate();
    UnprotectResultsSheet(resultsSheet);  // whether legacy/temporary
    let [rangeFreeze,rangeChange] = GetProtectResultsRanges(maxChangeROWS);
    ReprotectResultsRanges(rangeFreeze,rangeChange);
    if (lc.debug)
      Logger.log("Reprotections applied on results sheet, "+runnerNameId);
  }
}

/** 
 * Reprotects results sheet for each runner, allowing permitted users
 * to add/change latest results.
 */
function ReprotectResultsSheetPerRunner() {   // entry-point usage
  const runnersRANGE = runnerNameCOLUMN+lc.resultsStartROW+":"+runnerSurnameCOLUMN;
  let runnerFullNames = lv.allRunnersSheet.getRange(runnersRANGE)
    .getValues()
    .filter(name => name[0] != "")
    .map(name => [name[0],name[1]]);
  runnerFullNames.forEach((runnerFullName,index) => {
      var runnerName = runnerFullName[0];
      var runnerNameId = runnerName+'_'+index;
      ReprotectResultsSheet(runnerNameId);
    }
  );
  Logger.log("Reprotection completed on each results sheet");
}

//------------------------------------------------------------------------------------------

const editCOLUMN = "N";   // on the RHS of the results (at the bottom)
const onEditCOL = editCOLUMN.charCodeAt(0)-64 ;
const activeVALUE = '...entering';
const activeNOTE = 'If available, please include a HYPERLINK formula to any official result '+
  '(behind the date of the event) and enter your age-based grading (after calculating) together '+
  'with your gender position, before entering DONE here (ready for protection to be re-applied)';

/**
 * Appends a new result row to a runner's results sheet (without degrading MAP formulae)
 *    @param {GoogleAppsScript.Spreadsheet.Sheet} resultsSheet - for appending blank row
 */
function AppendNewResultRow(resultsSheet) {
  resultsSheet.appendRow([""]);
  let lastResultRow = resultsSheet.getLastRow();
  resultsSheet.getRange(lastResultRow,lc.PBtickBoxCOL)
    .clearContent();  // ensure undefined (instead of FALSE) for MAP to be effective
  resultsSheet.setActiveSelection('A'+lastResultRow);
  return lastResultRow;
}

/**
 *  Appends a non-parkrun result row to the active results sheet if user is permitted
 *  to allow them to add their result manually and "return" via an onEdit function 
 */
function AppendNonParkrunResult() {     // user 
  const resultsSheet = SpreadsheetApp.getActiveSheet();   // Applies to active (results) sheet...
  const runnerNameId = resultsSheet.getName();
  let [runnerName,index] = runnerNameId.split('_');
  if (isNaN(index)) {
    Logger.log('ERROR: '+runnerNameId+' is NOT a results sheet (like Alan_13)');
    return;
  } else if (lc.debug)
    Logger.log('Attempting to add non-parkrun result to '+runnerNameId+'...');
  let userEmail = Session.getEffectiveUser().getEmail();  // user permitted?
  let protection = resultsSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE)[1];
  // const ownerEmail = activeSpreadsheet.getOwner().getEmail();
  let editorEmails = (protection)   // get actual list of editors from existing protection
    ? protection.getEditors().map(editor => editor.getEmail())
    : GetResultsEditorsForRunner(runnerNameId);   // ...OR get expected list (if not already set)
   if (lc.debug)
    Logger.log('Checking user, '+userEmail+' is in the authorised list:\n'+
      editorEmails+'\n'+', which includes the spreadsheet owner');
  if (editorEmails && editorEmails.includes(userEmail)) { 
    let lastResultRow = AppendNewResultRow(resultsSheet);
    resultsSheet.getRange(editCOLUMN+lastResultRow)    // N<lastRow>
      .setValue(activeVALUE)
      .setFontStyle('italic')
      .setNote(activeNOTE);
    resultsSheet.getRange(lastResultRow,1).activate();
    // CAUTION: An Editor (non-Owner) is unable to re-apply protections (relies on owner's onEdit trigger)
    //  ReprotectResultsSheet(runnerNameId);
  } else
    Logger.log('ERROR: Runner has not been granted permission to add/change results in sheet, '+runnerNameId);
}

/**
 * This need to reprotect is triggered (for the Owner), typically after AppendNonParkrunResult (by User)
 * in order to reprotect their results sheet.
 *    @param {Object} event - Google Sheets onEdit trigger event
 */
function onEditDetectNeedToReprotect(event) {  // trigger on Edit runs as spreadsheet owner
  const col = event.range.getColumn();
  if (lc.debug)
    Logger.log('Column: '+col+'; Row: '+event.range.getRow()+'; Value: '+event.value);
  if (col !== onEditCOL || event.value === activeVALUE)
    return;   // exit ASAP to reduce interruption from interim/other edits
  let resultsSheet = event.range.getSheet();
  // Skip sheet if table does not contain Event results
  if (resultsSheet.getRange(lc.eventHeaderCELL).getValue() !== lc.resultTABLE) return;
  let runnerNameId = event.source.getName();    // OR resultsSheet.getName()
  const resultRow = event.range.getRow();
  if (resultsSheet.getRange(resultRow,2).getDisplayValue()) {
    if (lc.debug)
      Logger.log('Reprotecting results sheet, '+runnerNameId+'...');
    resultsSheet.activate();      // pre-requisite to get range
    ReprotectResultsSheet(runnerNameId);  // shift protection down 1 row
    event.range.clearContent();  // Reset FULLY inactive when DONE with latest 
  } else {
    resultsSheet.deleteRow(resultRow);
    Logger.log('WARNING: Deleted empty/incomplete result appended to sheet, '+runnerNameId+
      '; entry of date and time is essential');
  }
}

function ScrollBeyondLastResult() {
  const functionNAME = "ScrollBeyondLastResult";
  // Select the first empty row in preparation for pasting a non-parkrun entry
  // Conveniently used to skip to the bottom of the table below the latest result
  var resultsSheet = SpreadsheetApp.getActiveSheet();
  if (lc.debug)
    Logger.log('Running '+functionNAME+' on sheet: '+resultsSheet.getName());
  var lastRow = resultsSheet.getLastRow();
  if (lastRow <= lc.resultsStartROW) return; // table has no results yet
  // Find an empty row from the bottom of the table since more efficient
  var ri = lastRow;
  for (; ri>lc.resultsStartROW; ri--) 
    if (resultsSheet.getRange(ri,lc.pasteCOL).getDisplayValue()) break;
  // When last row has a result, prepare for pasting new result(s) below
  if (ri == lastRow) {
    var newRow = resultsSheet.appendRow([""]);
    resultsSheet.getRange(newRow.getLastRow(),lc.PBtickBoxCOL).clearContent();
  }
  // Always select the row below the last result
  resultsSheet.getRange(ri+1,lc.pasteCOL).activate();
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
  var range = e.range;
  // Only effective when selecting a cell just to the right of the table
  // assuming a down-arrow image is on the right of that Event result header
  if (range.getColumn() !== scrollCOL) return
  var resultsSheet = e.source.getActiveSheet();
  // Skip sheet if table does not contain Event results
  if (resultsSheet.getRange(lc.eventHeaderCELL).getValue() !== lc.resultTABLE) return;
  // Skip also if first result has not been entered (and cleaned manually)
  if (resultsSheet.getRange(lc.firstResultCELL).getValue() === "") return;
  ScrollBeyondLastResult();
}

/* ---------------------------------------------------------------------------
/
/   The following definitions and functions are used to support the automatic
/   creation of comparison charts.
/   The SpreadSheet itself has built-in functions to support automated charts.
/   These include standard comparative graphs for different groups that may be
/   customised to suit what best inspires a family or club.
/   There are eight samples of filtered charts (mostly by Age-Grade %age):
      Juniors (over past year)
      Seniors (over past year)
      Veterans (under 50, over past year)
      Supervets (50+,over past year)
      By FAMILY surname (overpast year)
      Females (by Gender position, over past year)
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
                INDEX(fResult,fastTime,ageCat),
                idx
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

    Thereafter, each challenge section may be hidden, when a new section is added with a new planned date.

    For maintenance needs, here is the hierarchy of functions for the comparative charts:

      GenerateChartsFromGroups (four Group sheets)
        PrepassBlocksOfGroups
        For each block of groups...
          For each group required, related to same target sheet...
            ClearGroupChartsOnSheet (e.g. Age Groups, unless continuing)
            ExtractGroupRunners (per group)
            GenerateGroupChartInSheet (if enough time)
              FilterDatedGroupResults
                CollateDatedGroupResults
              CopyGroupResultsToSheet
              EmbedGroupResultsChart
                ApplyFormatsOnGroupResultsChart
              If max time would be exceeded for next chart
                trigger -> GenerateChartsFromGroups (to continue same block)

      GenerateAgeGroupCharts    
        GenerateChartsFromGroups (Age-Groups)
      GenerateLeagueGroupCharts    
        GenerateChartsFromGroups (Leagues)
      GenerateFamiliesGroupCharts    
        GenerateChartsFromGroups (Families)
      GenerateGenderGroupCharts    
        GenerateChartsFromGroups (Gender)

      GenBatchChartsFromGroups (in 4 batches)
        triggers -> GenerateAgeGroupCharts
        triggers -> GenerateGenderGroupCharts (in parallel)
        triggers -> GenerateLeagueGroupCharts (after 10 minutes)
        triggers -> GenerateFamiliesGroupCharts (in parallel)

      When any change of hex colour codes in the top Legends table...
        ColourLegendsInGroups
/
/ ---------------------------------------------------------------------------
*/

/**
 * Returns an array of selective runners performances versus event dates,
 *  ensuring event dates are unique (even if event location differs),
 *  based on most recent years (unless all performances override).
 *    @param {Array<Date>} filteredDates  - Array of non-unique event dates
 *    @param {Array<string>} runners      - 2D Array of selective runners' names with unique indices
 *    @param {Object} runnersPerfs        - Object of performances for each indexed runner on dates (subset of all dates)
 *    @param {number} [mostRecentYears=recentYRS] - Number of recent years to include (0 to get all)
 *  @returns {Array<Array>} as an Array of arrays containing dated performances (with nulls for absences)
 *           (potentially more efficient as a 2D array)
 */
function CollateDatedGroupResults(filteredDates,runners,runnersPerfs) {  // pre-filtered
  let uniqueDates = [...new Set(filteredDates.map( // remove duplicates and return as Array
    date => new Date(date).getTime()))]
      .map(timestamp => new Date(timestamp))  // use timestamps to be sure (assumimg midnight)
      .sort((a,b) => a-b)                     // sort first and then return to be string
      .map(date => Utilities.formatDate(date,
        Session.getScriptTimeZone(),lc.dateFORMAT));
  return uniqueDates.map(date => [
    date,
    ...runners.map(runner => {
      let runnerNameId = runner.join('_');
      return runnersPerfs[runnerNameId]?.[date] || null;  
    })
  ]);
}

/**
 * Retrieves and collates dated performances for the specified runners from their results sheets.
 * This filters by date BEFORE collating instead of filtering during the collation of results
 *    @param {string}                                - chart title,only used for tracking distinct groups
 *    @param {Array<string>}                         - runners - Array of runner names & indices from Group
 *    @param {number} [mostRecentYears=lc.recentYRS] - Number of recent years to include (0 to get all)
 *    @param {number} [perfIndex=lc.ageGradeINDEX]  - Column index of performance values (Age Grade, Time, etc.)
 *  @returns {Array<Array>} Array of arrays containing dated performances (with nulls for absences)
 */
function FilterDatedGroupResults(chartTitle,runners,
  mostRecentYears = lc.recentYRS,   // assume all performances if zero years cut-off
  perfIndex = lc.ageGradeINDEX)   // default presents values <1 as %ages
{
  const NOW = new Date();
  let cutoffDate = mostRecentYears == 0 
  ? new Date('2000-01-01') 
  : new Date(
      NOW.getFullYear() - (mostRecentYears|0), 
      NOW.getMonth() - ((mostRecentYears%1)*12|0), 
      NOW.getDate()
    );
  var filteredDates = [];
  var runnersPerfs = {};
  for (var mi=0; mi<runners.length; mi++) {
    runner = runners[mi];
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
          SpreadsheetApp.ButtonSet.OK);
        return;
      }
      let numResults = resultsSheet.getLastRow()-lc.resultsStartROW+1;
      let resultsRange = resultsSheet.getDataRange()
        .offset(lc.resultsStartROW-1,0,numResults);
      var startResult = -1;    // to find first result from the cutoff date
      let indexedDates = resultsRange.offset(0,lc.dateINDEX,numResults,1)
        .getValues()      // for dates only to filter
        .filter((date,result) => {
          if (startResult >= 0)
            return true
          if (new Date(date[0]) > cutoffDate) {
            startResult = result;
            return true;
          } else return false;
        });       // reversing before and after is slower
      let numFilteredResults = indexedDates.length;   // = numResults-startResult;
      // Logger.log('2b1A-start. Start dated result, '+startResult+ ' for runner,'+runnerName);
      if (startResult >= 0) {   // 0 if runner's very first result after cut-off date
        let perfs = resultsRange.offset(startResult,perfIndex,numFilteredResults,1)
          [perfIndex == lc.timeINDEX
                        ? 'getDisplayValues'
                        : 'getValues']
          ().map((result,mi) => resultsSheet.isRowHiddenByUser(lc.resultsStartROW+startResult+mi)
            ? null    // although date retained for alignment, blankout if distorts chart
            : result[0]);
        runnersPerfs[runnerNameId] = {};
        // Logger.log('2b1A-shown. Gathered visible results for runner,'+runnerName);
        for (var di=0; di<indexedDates.length; di++)
          runnersPerfs[runnerNameId][indexedDates[di]] = perfs[di];
        // Logger.log('2b1A-for. Recorded results for runner,'+runnerName);
        filteredDates = filteredDates.concat(indexedDates);
      }
    } catch (err) {
      Logger.log("ERROR: filtering results for unique runner, " +runnerName+'['+runnerIndex+"]\n"+err);
    }
  }
  if (lc.debug)
    Logger.log('Filtered Dated Group Results for all: '+runners);
  runnersDatedPerfs = CollateDatedGroupResults(
    filteredDates,runners,runnersPerfs);
 if (lc.debug)
    Logger.log('Collated Dated Group Results for all: '+runners);
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
 *    @param {number} [groupPerfsCol=lc.groupPerfsCOL]         - Group Column in each Performances sheet
 *    @param {string} [dateFormat=lc.dateFORMAT]   - Format used for displayed dates
 *    @param {number} [decimalPlaces=4]         - Floating precision (unless integer)
 * @returns {Range} where performances are on the sheet (beyond subsequent charts)
 */
function CopyGroupResultsToSheet(
  perfsSheet,runners,runnersPerfs,
  groupPerfsRow = 2,
  groupPerfsCol = lc.groupPerfsCOL,
  dateFormat = lc.dateFORMAT,     // aligned  with display format
  decimalPlaces = 4)
{
  var dates = ['Date', ...runnersPerfs.map(date => date[0])];
  var groupTable = [dates];   // 'Date' with d-MMM-YY dates are headers
  for (var ci=1; ci<=runners.length; ci++) {
    var runnerPerfs = runnersPerfs.map(result => result[ci]);
    var [runnerName,runnerIndex] = runners[ci-1];
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
  groupPerfsRange.offset(0,1,1,groupTable[0].length-1)  // exclude 1st column (containing 'Date')
    .setNumberFormat(dateFormat);    // sets dates format as expected displayed on results sheets
  var colsToClear = perfsSheet.getLastColumn()-(groupPerfsCol+groupTable[0].length-1);
  if (colsToClear>0)
    groupPerfsRange.offset(0,groupTable[0].length,  // clear any surplus from a previous run...
      groupTable.length,colsToClear)                // ...where earliest event dates lapse (fewer results)
        .clearContent();                            // ...OR if cut-off years reduced results considered
  if (lc.debug)
    Logger.log('Formatted dates are:\n'
      +groupPerfsRange.offset(0,1,1,groupTable[0].length-1).getDisplayValues());
  return groupPerfsRange;
}

/**
 * Apply format to group performances and prepare space on sheet for the chart
 *     @param {Sheet}  perfSheet       - The sheet where spacing & formatting applies
 *     @param {array}  runnersLegend   - The runners in the legend (to be coloured)
 *     @param {Range}  groupPerfsRange - The range of performances (with Date header)
 *     @param {number} perfChartRow    - The row where the chart will be placed
 *     @param {number} perfChartCol    - The column where the chart will be placed
 *     @param {string} perfFormat      - The format to apply to performance values
 *     @param {number} chartHeight     - The height of the chart
 *     @param {number} chartWidth      - The width of the chart
 *     @param {number} offsetBorder    - The offset border size (in pixels)
 */
function ApplyFormatsOnGroupResultsChart (
  perfSheet,groupPerfsRange,runnersLegend,
  perfChartRow,perfChartCol,perfFormat,
  chartHeight,chartWidth,offsetBorder=lc.offsetBORDER)
{
  const padFACTOR = 0.05;       // adjust min/max values to encourage "growth"
  const cellHeightSIZE = 21;    // cell height accounts for merged cells on chart
  const surroundFACTOR = 2;     // pad out border all around the chart (in pixels)
  const rotateDownANGLE = -90;  // ensures Dates header text appears downward
  perfSheet.setRowHeight(perfChartRow,chartHeight+surroundFACTOR*offsetBorder
    -cellHeightSIZE*runnersLegend.length);  // Adjust for merging rows (later)
  perfSheet.setColumnWidth(perfChartCol,chartWidth+surroundFACTOR*offsetBorder);
  if (lc.debug)
    Logger.log('Rows: '+groupPerfsRange.getNumRows()+', '+'Cols: '+groupPerfsRange.getNumColumns());
  perfsRange = groupPerfsRange.offset(1,1,  // Get runners' performances only...
    groupPerfsRange.getNumRows()-1,         // ...ignoring the Date row
    Math.max(1,groupPerfsRange.getNumColumns()-1));     // ...and the runner name column
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
  return {min: minPerf, max: maxPerf};
}

/**
 * Embeds a line chart of specified groups of runners' performances in an existing sheet.
 *  Prior to creating a Group chart for comparing related runners, it is assumed that
 *  the performances have already been collated from each of the grouped runners 
 *  results and have already been presented to the right of where this Group chart is
 *  to be created on this same sheet.
 *    @param perfSheet         - The sheet (object) to embed the chart in
 *    @param {string} chartTitle    - The title of the chart
 *    @param {range} groupPerfsRange      - Results range on RHS of chart position
 *    @param {Array<Array<string>>} runnersLegend - A 2D array of names, indices & colours.
 *    @param {number} [perfChartRow=2]    - The row number to place the chart at
 *                                           (avoid subsequent charts coinciding)
 *    @param {number} [perfChartCol=2]    - The column number to place the chart 
 *                                           (typically starting at B2)
 *    @param {boolean} [showDates=false]  - Allow non-contiguous dates on the horizontal
 *    @param {boolean} [reverseTrend=1]   - Trends in reverse: rising to low values
 *    @param {boolean} [stripIndex=false] - Simplify legend if first name context unique
 *    @param {number} [filterRecentYears=lc.recentYRS] - The cut-off years for results 
 *    @param {string} [perfTitle=lc.chartYAxisTITLE]   - Performance title on vertical
 *                                                    (typically Age Grade in results)
 *    @param {string} [perfFormat=lc.ageGradeFORMAT]   - Format the performance (e.g. %age)
 */
function EmbedGroupResultsChart(
  perfSheet,chartTitle,groupPerfsRange,runnersLegend,
  perfChartRow = 2,
  perfChartCol = 2,
  showDates=false,reverseTrend=1,stripIndex=false,
  filterRecentYears = lc.recentYRS,
  perfTitle = lc.chartYAxisTITLE,
  perfFormat = lc.ageGradeFORMAT)
{
  const chartWIDTH = 800;
  const chartHEIGHT = 350;
  var perfLimits = ApplyFormatsOnGroupResultsChart(
    perfSheet,groupPerfsRange,runnersLegend,
    perfChartRow,perfChartCol,perfFormat,
    chartHEIGHT,chartWIDTH,lc.offsetBORDER);
  if (filterRecentYears === 1)
    chartTitle += ' (past year)';
  else if (filterRecentYears > 0)
    chartTitle += ' (max '+filterRecentYears+' years)';
  else  // no date filter
    chartTitle += ' (best ever)';
  perfSheet.getRange(perfChartRow-1,perfChartCol)
    .setValue(chartTitle)   // repeat the chart title in the row above...
    .setFontSize(14)        // ...highlighted above chart (assuming not disabled)
    .setFontWeight('bold');
  if (stripIndex) {  // tweak to remove the unique index if desired on chart legend
    let numRunners = groupPerfsRange.getHeight()-1;
    let namesRange = groupPerfsRange.offset(1,0,numRunners,1);
    namesRange.offset(0,-1,numRunners,1)  // before stripping, preserve index in preceding column... 
      .setValues(namesRange.getValues()   // ...to get unique matching selection for App users
        .map(name => [name[0].split('_')[1]]));
    namesRange
      .setValues(namesRange.getValues()
        .map(name => [name[0].split('_')[0]]));
  }
  let seriesOptions = Object.assign({},
    ...runnersLegend.map((colour,i) => ({
      [i]: {color: colour[2]}   // 3rd column of legend
    }))
  );
  var embeddedChart = perfSheet.newChart()
    .asLineChart()
    .addRange(groupPerfsRange)
    // .setMergeStrategy(Charts.ChartMergeStrategy.MERGE_ROWS)  // if multiple sheet ranges?
    // .setHiddenDimensionStrategy(Charts.ChartHiddenDimensionStrategy.IGNORE_BOTH) // always shows
    .setTransposeRowsAndColumns(true)         // data on same sheet in D column?
    .setNumHeaders(1)
    .setPosition(perfChartRow,perfChartCol,lc.offsetBORDER,lc.offsetBORDER)
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
 * Retrieves the Group Results sheet and clears any charts from it.
 *  A blank sheet is created if the Group results sheet does not exit
 *    @param {string} [perfSheetName="Leagues"] - The name of the sheet to retrieve and clear.
 *  @returns {Spreadsheet.Sheet} - the new/cleared sheet object.
 */
function ClearGroupChartsOnSheet(
  perfSheetName = "Leagues",
  index = 0)
{
  var perfSheet = activeSpreadsheet.getSheetByName(perfSheetName);
  if (!perfSheet)
    try {
      perfSheet = activeSpreadsheet.insertSheet(perfSheetName);
    } catch (err) {
      return null;
    }
  if (index === 0) {  // fresh start from the top?
    var charts = perfSheet.getCharts();
    var numCharts = charts.length;
    if (numCharts > 0) {
      var charts = perfSheet.getCharts();
      for (var i = charts.length-1; i >= 0; i--)
        perfSheet.removeChart(charts[i]);
    }
    perfSheet.getDataRange().breakApart();  // remove merges to allow shift between tables
    // Clear content, styles, formats, backgrounds, and text rotations
    //    retain same number of rows and columns
    var maxRows = perfSheet.getMaxRows();
    var maxCols = perfSheet.getMaxColumns();
    perfSheet.getRange(1, 1, maxRows, maxCols).clear({
      contentsOnly: false, // Wipes the values/formulas
      formatOnly: false,   // Wipes fonts, alignments, fills, text direction
      validationsOnly: false, 
      commentsOnly: false
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
  for (var gj=0; gj<runnersNames.length; gj++) {
    if (runnersNames[gj] == "")
      break;
    runnersLegend.push([runnersNames[gj],runnersIndices[gj],runnersColours[gj]]);
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
function GenerateGroupChartInSheet(
  perfSheet,chartTitle,runnersLegend,
  perfChartRow,perfChartCol,
  showDates=false,reverseTrend=1,stripIndex=false,
  filterRecentYears,
  perfColumnTitle=lc.chartYAxisTITLE,perfColumnIndex,perfFormat)
{
  if (!perfSheet) return null;
  let runners = runnersLegend.map(runner=>[runner[0],runner[1]]); // include unique Id
  // runners are a subset! so index is EITHER from Runners sheet OR in Legend
  var runnersDatedPerfs = FilterDatedGroupResults(
    chartTitle,runners,filterRecentYears,perfColumnIndex);
  if (!runnersDatedPerfs) return null;
  var groupPerfsRange = CopyGroupResultsToSheet(perfSheet,runners,
    runnersDatedPerfs,perfChartRow);
  if (!groupPerfsRange) return null;
  var perfChart = EmbedGroupResultsChart(
    perfSheet,chartTitle,groupPerfsRange,runnersLegend,
    perfChartRow,perfChartCol,
    showDates,reverseTrend,stripIndex,
    filterRecentYears,
    perfColumnTitle,perfFormat);
  if (perfChart) {
    Logger.log('Generated group chart, '+chartTitle+
      ' in row, '+perfChartRow+' of sheet, '+perfSheet.getName()+
      ' ('+runners.length+' runners)');
  }
}

var defaultPerfSheetName = 'Age Groups';    // may vary if overridden

/**
 * Builds block (array) of group params (arrays) from Groups sheet configuration..
 * Assume null for parameter (if empty) - see notes on Groups sheet
 *    @param {Sheet} groupsSheet - e.g. for Groups sheet in GAS spreadsheet 
 *  @returns {Object} grpBlocks - with perfs sheet names as keys, each with an array of group params
 * Note: a blank perfs sheet name (in col A) implies that of the group above
 *  - initially 'Age Groups' if col A completely empty (not recommended; not so in template)
 */
function PrepassBlocksOfGroups(groupsSheet,perfSheetNames,groupsCount) {
  var grpBlocks = {};
  for (var gi=0; gi<groupsCount; gi++) {
    var group1stRow = lc.groupsStartROW+lc.numGroupROWS*gi;
    var params = groupsSheet.getRange(group1stRow,1,1,lc.groupStringsCOUNT)
      .getDisplayValues()[0]
      .concat(
        groupsSheet.getRange(group1stRow,lc.groupStringsCOUNT+1,1,lc.groupValuesCOUNT)
          .getValues()[0]
      );
    let perfSheetName = params[0] || defaultPerfSheetName;
    defaultPerfSheetName = perfSheetName;
    if (!perfSheetNames.includes(perfSheetName))
      continue;   // skip groups in this block because out of scope
    if (!grpBlocks[perfSheetName])
      	grpBlocks[perfSheetName] = [];
    var grpParams = {};
      for (var pi=0; pi<lc.paramsCOUNT; pi++) {
        switch (pi) {    // from columns B..K in 1st row
          case 0: grpParams.sheetName = perfSheetName; break;   // for completeness
          case 1: grpParams.groupName = params[1]; break;
          case 2: grpParams.groupTitle = params[2]; break;
          case 3: grpParams.perfColumnTitle = params[3] || lc.chartYAxisTITLE; break;
          case 4: grpParams.perfFormat = params[4] || undefined; break;
          case 5: grpParams.perfColumnIndex = params[5] || undefined; break;
          case 6: grpParams.filterRecentYears = params[6] || undefined; break;
          case 7: grpParams.showDates = (params[7] == true); break;
          case 8: grpParams.reverseTrend = (params[8] == true) ? -1 : 1; break;
          case 9: grpParams.stripIndex = (params[9] == true); break;
          case 10: grpParams.disableDraw = (params[10] == true); break;
        }
      }
    let maxRunnersCount = groupsSheet.getLastColumn()-lc.runnersStartCOL+1;
    grpParams.runnersRange = groupsSheet.getRange(
      group1stRow+1,lc.runnersStartCOL,
      lc.numRunnerROWS,maxRunnersCount);
    grpBlocks[perfSheetName].push(grpParams);
  }
  return grpBlocks;
}

function getCaller() {
  let stack = new Error().stack.split('\n');
  return stack.length > 3
    ? stack[3].match(/at (\w+)/)[1]
    : undefined;
}

/**
 * Cleans up triggers after any for generating charts
 *  @param {string} [classScript='Generate'] - starts like this?
 */
function CleanupCoreBatch(
  classScript = 'Generate')
{
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = triggers.length-1; i >= 0; i--) {
    var trigger = triggers[i];
    let foundTrigger = trigger.getHandlerFunction();
    Logger.log('Trigger handler cleaning: '+foundTrigger);
    if (trigger && foundTrigger.indexOf(classScript) === 0)
      ScriptApp.deleteTrigger(trigger);
  }
}

/**
 * Generates charts in performance sheets based on config in the Groups sheet.
 * Groups are organised into blocks, each with a different target sheet - e.g Leagues
 *    @param {string} [groupsSheetName="Groups"] - existing sheet with Groups config
 */
function GenerateChartsFromGroups(
  groupsSheetName = 'Groups',
  perfSheetNames = ['Age Groups','Leagues','Families','Gender'])
  // perfSheetNames = ['Age Groups'])    // test performance
{
  const groupsSheet = activeSpreadsheet.getSheetByName(groupsSheetName);
  if (!groupsSheet) return;
  let groupsEndRow = groupsSheet.getLastRow();
  let groupsCount = (groupsEndRow-lc.groupsStartROW+1)/lc.numGroupROWS;

  // Stage 1: Identify blocks of groups (per sheet output) with parameters from Groups sheet
  let grpBlocks = PrepassBlocksOfGroups(groupsSheet,perfSheetNames,groupsCount);
  if (lc.debug)
    Logger.log('Pre-pass '+groupsCount+' groups of charts for four separate target sheets');
  for (var perfSheetName in grpBlocks) {
    if (!perfSheetNames.includes(perfSheetName))
      continue;   // split scope into separate tasks: two blocks per scope
    if (lc.debug)
      Logger.log('Perf sheet: '+perfSheetName+' ['+grpBlocks[perfSheetName].length+' charts]');
    if (grpBlocks[perfSheetName].every(g => g.disableDraw))
      continue;
    let continueGenerate = getCaller();   // re-invoke if generation gets close to max time limit
    var index = parseInt(PropertiesService.getScriptProperties()
      .getProperty('index_' + perfSheetName)) || 0;   // where to continue if re-invoked
    // ONLY clear/clean a perf chart sheet if one or more charts NOT disabled
    var perfSheet = ClearGroupChartsOnSheet(perfSheetName,index);
    if (lc.debug && index === 0)
      Logger.log('Cleared all charts from sheet, '+perfSheetName);
    for (; index<grpBlocks[perfSheetName].length; index++) {
      var grpParams = grpBlocks[perfSheetName][index];
      if (lc.debug)
        Logger.log('Group '+index+': '+JSON.stringify(grpParams));

      // Stage 2a. Establish valid group of runners with colour legend...
      if (grpParams.disableDraw || !grpParams.groupName)
        continue;
      let runnersLegend = ExtractGroupRunners(grpParams.runnersRange);
      if (!runnersLegend || runnersLegend.length == 0)
        continue;
      let groupRange = groupsSheet.getRange(grpParams.runnersRange.getRow()-1,1);
      let perfChartRow = groupRange.offset(1,0).getValue() || undefined;
      let perfChartCol = groupRange.offset(1,1).getValue() || undefined;
      let firstRunner = groupRange.offset(1,2).getValue() || undefined;
      if (firstRunner === "#N/A")
        continue;

      // Stage 2b. Place Group performances data on sheet prior to creating Group chart
      var chartTitle = grpParams.groupName+" ("+grpParams.groupTitle+") "+grpParams.perfColumnTitle;
      GenerateGroupChartInSheet(
        perfSheet,chartTitle,runnersLegend,
        perfChartRow,perfChartCol,
        grpParams.showDates,grpParams.reverseTrend,grpParams.stripIndex,
        grpParams.filterRecentYears,grpParams.perfColumnTitle,
        grpParams.perfColumnIndex,grpParams.perfFormat
      );
      if (index+1 == grpBlocks[perfSheetName].length)
        continue;
      let remainingTimeSecs = lc.maxTimeSECS-Math.floor((new Date().getTime()-lv.startTime)/1000); 
      if (lc.debug)
        Logger.log('Script time remaining: '+remainingTimeSecs+' seconds');
      if (remainingTimeSecs < lc.chartTimeSECS) {
        Logger.log('WARNING: Insufficient time ('+remainingTimeSecs+' secs) to generate any more charts.');
        if (continueGenerate) {
          Logger.log('Exiting loop and continuing chart generation beyond index, '+index+
            ' by re-triggering '+continueGenerate);
          PropertiesService.getScriptProperties()
            .setProperty('index_'+perfSheetName,index+1);
          ScriptApp.newTrigger(continueGenerate)
            .timeBased()
            .after(10*1000) // resume in 10 seconds
            .create();
        } else  // need to use split functions for a large club/family
          Logger.log('ERROR: Exiting since unable to continue chart generation beyond index, '+index+
           ' without using the predefined split functions (such as GenerateAgeGroupCharts)');
        return;
      }
    }   // Generated all charts for this perfSheetName 
    PropertiesService.getScriptProperties()
          .deleteProperty('index_' + perfSheetName);
  }
}
               
function GenerateAgeGroupCharts() {
  GenerateChartsFromGroups('Groups',['Age Groups']);
}
function GenerateLeagueGroupCharts() {
  GenerateChartsFromGroups('Groups',['Leagues']);
}
function GenerateGenderGroupCharts() {
  GenerateChartsFromGroups('Groups',['Gender']);
}
function GenerateFamiliesGroupCharts() {
  GenerateChartsFromGroups('Groups',['Families']);
}

function GenBatchChartsFromGroups() {
  const batchGapMINS = 20;
  const parallelGapMINS = 9;
  CleanupCoreBatch("Generate"); // clear to go; avoid max triggers
  ScriptApp.newTrigger('GenerateAgeGroupCharts')
    .timeBased()
    .after(5000)
    .create();
  ScriptApp.newTrigger('GenerateGenderGroupCharts')
    .timeBased()
    .after(parallelGapMINS*60000)  // 9 minutes later in parallel
    .create();
  ScriptApp.newTrigger('GenerateLeagueGroupCharts')
    .timeBased()
    .after(batchGapMINS*60000)  // after 20 mins delay (after 3 consecutive Age-group runs)
    .create();
  ScriptApp.newTrigger('GenerateFamiliesGroupCharts')
    .timeBased()
    .after((batchGapMINS+parallelGapMINS)*60000)  // 9 minutes later in parallel
    .create();
}

function ColourLegendsInGroups(
  sheetName="Groups")
{
  var sheet = activeSpreadsheet.getSheetByName(sheetName);
  var dataRange = sheet.getDataRange();
  var values = dataRange.getValues();
  for (var gi=0; gi<values.length; gi++) {
    if (values[gi][1] === "Legend") {
      for (var lj=1; lj<values[gi].length; lj++) {
        var hexValue = values[gi][lj];
        if (typeof hexValue === "string"
            && hexValue.match(/^#[0-9A-F]{6}$/gi))
        {
          sheet.getRange(gi+1,lj+1).setBackground(hexValue);
          sheet.getRange(gi+1,lj+1).setFontColor("#ffffff");  // TODO or #000000 for contrast?
        }
      }
    }
  }
}

function GetRelatedTabColor(
  nameIndex = 'Alan_13')
{
  let resultsSheet = lv.activeSpreadsheet.getSheetByName(nameIndex);
  const colour = resultsSheet.getTabColor();
  Logger.log('Colour: '+colour);
  return colour || "#ffffff"; // default if no color
}

/* ---------------------------------------------------------------------------
/
/   The following definitions and functions support the generation of the Appsheet tables for the mobile App user
/   There are two use-case heirarchies of functions:
/
/     1.  PrepareAppSheets (for all users?)
/           GetOverview 
/           GetPerfCharts
//          for each Perf sheet block (e.g. Age Groups)...
//            for each chart in group block....
/               GetChartRange
/               GetChartRunnersNames
/
/     2.  GetRunnerTrendCharts (for a user)
//        for each user chart (on user result's sheet)...
/           GetChartRange
/
/  ---------------------------------------------------------------------------
*/

function GetChartRunnersNames(resultsRange) {
  let numRunners = resultsRange.getNumRows()-1;
  return resultsRange.offset(1,-1,numRunners,2)  // e.g. D2:O4 => C3:D4
    .getValues().flat();
}

function GetChartRange(
  resultsRange,
  chartRowOffset = 0,
  chartColOffset = lc.chartColOFFSET,
  numRows = 1,
  numCols = 1)
{
  return resultsRange.offset(chartRowOffset,chartColOffset,
    numRows,numCols);      // e.g. D2:O4 => B2:B4
}

function GetChart(sheetName,chartId) {
  var sheet = lv.activeSpreadsheet.getSheetByName(sheetName);
  var chart = sheet.getCharts().find(c => c.getId() === chartId);
  if (chart) {
    Logger.log('Title: ' + chart.getOptions().get('title'));
    Logger.log('Range: ' + chart.getRanges()[0].getA1Notation());
  }
  return chart;
}

/**
 * This updates the Oveview sheet content after extracting the owner, spreadsheet Id and results sheet gids
 */
function GetOverview() {
  var overviewTable=[["Overview","Value"]];
  overviewTable.push(["Spreadsheet",lv.activeSpreadsheetName]);
  overviewTable.push(["SS Id",lv.activeSpreadsheetId]);
  let owner = lv.activeSpreadsheet.getOwner().getEmail();
  overviewTable.push(["Owner",owner]);
  const resultsSheets = [
    ["Runners","Runners Gid"],
    ["Latest","Latest Gid"],
    ["Rankings","Rankings Gid"],
    ["Challenges","Challenges Gid"]
  ];
  // "#" is derived from Runners! - assume =COUNTIF(Runners!K:K,TRUE)
  // "Parkrun Group No." is in Runners! - assume J1
  // "Parkrun Domain" is in Runners! - assume D1
  var resultsSheetIds = [];
  resultsSheets.forEach(([resultsSheetName,resultsSheetIdHeader]) => {
    let resultsSheet = lv.activeSpreadsheet.getSheetByName(resultsSheetName);
    let resultsSheetId = resultsSheet.getSheetId();
    overviewTable.push([resultsSheetIdHeader,resultsSheetId]);
  });
  let csv = overviewTable.map(row => row.join(',')).join('\n');
  Logger.log(csv);
  let overviewSheet = lv.activeSpreadsheet.getSheetByName('Overview');
  // overviewSheet.clear();  // already exists in template and pre-formatted
  // let titleRange = overviewSheet.getRange(1,1,1,2);    // table header (below title) has 2 columns (for transposing)
  overviewSheet.getRange(1,1,overviewTable.length,overviewTable[0].length)
    .setValues(overviewTable);
}

/**
 * This allows selection from the performance Charts, typically those in which a runner appears.
 *    @param {Array of strings} perfSheetNames - list of sheets with comparative charts
 *    @param {string} runnerNameId - optional <first name>_<index>, e.g. Alan_13
 *  returns {Array of key pairs} selectPerfCharts - perf. charts of those with runner, otherwise all are options 
 */
function GetPerformanceCharts(
  perfSheetNames = ['Age Groups','Leagues','Families','Gender'],
  runnerNameId)
  // runnerNameId ='Alan_13')   // for testing
{
  let selectPerfCharts = {};
  let chartsTable = [["Charts","=Overview!B$3","","","","","",""]];
  chartsTable.push(["Performance Chart","SS Id","Perf Gid","Perf Sheet","Range","Chart Id","#","Runners Ids"]);
  perfSheetNames.forEach(perfSheetName => {
    let perfSheet = lv.activeSpreadsheet.getSheetByName(perfSheetName);
    if (perfSheet) {
      perfSheetId = perfSheet.getSheetId();
      let perfCharts = perfSheet.getCharts();
      perfCharts.forEach(perfChart => {
        let chartId = perfChart.getChartId();
        let chartTitle = perfChart.getOptions().get('title');
        let resultsRange = perfChart.getRanges()[0];   // only one range exists
        // let resultsCells = resultsRange.getA1Notation();
        let numRunners = resultsRange.getNumRows()-1;
        let chartRange = GetChartRange(resultsRange,0,lc.chartColOFFSET,numRunners+1,1)
            .getA1Notation();
        if (lc.debug)
          Logger.log(chartTitle+'|'+chartId+'|'
            +perfSheetName+'|'+perfSheetId+'|'+chartRange);
        let runnersNames = GetChartRunnersNames(resultsRange);
        if (runnerNameId) {   // filter by runner
          if(runnersNames.some(indexedName => String(indexedName).includes(runnerNameId))   // select if any match indexed name
              || runnersNames.some((strippedName,i) => i % 2 === 1   // otherwise, index preserved and precedes name entry 
              && runnerNameId === strippedName+'_'+runnersNames[i-1])) {
            let resultsSheetId = lv.activeSpreadsheet
              .getSheetByName(runnerNameId)
              .getSheetId();
            selectPerfCharts[chartId] = {
              title: chartTitle,
              perfChartsheetId: perfSheetId,
              sheet: perfSheetName,
              range: chartRange,
              resultsSheetId: runnerNameId};
          }
        } else {  // file all for import
          let rowNum = chartsTable.length+1;
          let chartTitleLink =  // hyperlink needed to be added virtually inside AppSheet (but included here by example)
            '=HYPERLINK("https://docs.google.com/spreadsheets/d/"&Overview!B$3&"/view#gid="&C'+rowNum
            +'&"&range="&E'+rowNum+',"'+chartTitle+'")';
          let runnersIds = runnersNames.map((name, i) => {
            if (i % 2 === 1) return name.includes('_')
              ? name
              : name+'_'+runnersNames[i-1];
          }).filter(Boolean).join('|');
          chartsTable.push([chartTitleLink,"=B$1",perfSheetId,perfSheetName,chartRange,chartId,numRunners,runnersIds]);
        } 
      });
    }
  });
  if (runnerNameId) {
    // if (debug)
      Logger.log('Selective Results chart(s) for runner, '+runnerNameId+':\n'
        +JSON.stringify(selectPerfCharts));
    return selectPerfCharts;
  } else {
    let csv = chartsTable.map(row => row.join(',')).join('\n');
    Logger.log(csv);
    chartsSheet = lv.activeSpreadsheet.getSheetByName('Charts');
    // chartsSheet.clear();  // already exists in template and pre-formatted
    // let titleRange = chartsSheet.getRange(1,1,1,8);    // table header (below title) has 8 columns
    chartsSheet.getRange(1,1,chartsTable.length,chartsTable[0].length)
      .setValues(chartsTable);
  }
}

/**
 * Returns the number of runs completed by a runner
 * @param {string} runner ID
 * @return {number} of results
 */
function getNumRuns(
  runnerNameId='Alan_13')
{
  const numRunsCOL = 13;  // column M of Runners sheet
  let runnerIdx = Number(runnerNameId
    .split("_")[1]);
  let numRuns = lv.allRunnersSheet      //  assume current SS
    .getRange(lc.runnersStartROW+runnerIdx,numRunsCOL)
    .getValue();
  return numRuns ? numRuns : 0;
}

/**
 * Returns the sharing role for a given email on the active Spreadsheet
 * @param {string} email The email address to check
 * @return {string} "Owner", "Editor", "Viewer", or "Blocked"
 * @customfunction
 */
function getRole(
  runnerNameId='Alan_13')
{
  const emailCOL = 4;  // column D of Runners sheet
  const numRunsCOL = 4;  // column M of Runners sheet
  let runnerIdx = Number(runnerNameId
    .split("_")[1]);
  let email = lv.allRunnersSheet      //  assume current SS
    .getRange(lc.runnersStartROW+runnerIdx,emailCOL)
    .getValue();
  if (!email) return "Blocked";
  const ssFile = lv.activeSpreadsheet;
  if (ssFile.getOwner().getEmail() === email) return "Owner";
  const editors = ssFile.getEditors().map(u => u.getEmail());
  if (editors.includes(email)) return "Editor";
  const viewers = ssFile.getViewers().map(u => u.getEmail());
  if (viewers.includes(email)) return "Viewer"; // incl.  Commenter
  return "Blocked";
}

function GetSpreadsheetShares() {
  var ssFile = lv.activeSpreadsheet;
  const owner = ssFile.getOwner();
  const editors = ssFile.getEditors();
  const viewers = ssFile.getViewers();
  let output = [];
  output.push(["Type", "Email"]);
  output.push(["Owner", owner.getEmail()]);
  editors.forEach(u => output.push(["Editor", u.getEmail()]));
  viewers.forEach(u => output.push(["Viewer", u.getEmail()]));
  output.push(["",""]);
  const accessSheet = ssFile    // dump for easy summ,ary
    .insertSheet("Sharing_List_" + new Date().toISOString().slice(0,10));
  ssFile = DriveApp.getFileById(lv.activeSpreadsheetId);  // for more detail
  const access = ssFile.getSharingAccess(); // PUBLIC, DOMAIN, PRIVATE
  const permission = ssFile.getSharingPermission(); // VIEW, EDIT, NONE
  output.push(["Link Sharing", access.toString()]),
  output.push(["Permission", permission.toString()]);
  accessSheet.getRange(1,1,output.length, 2).setValues(output);
  Logger.log(output);
}

/**
 * This gets the results trend chart(s) from a runner's sheet.
 *    @param {string} runnerNameId - <first name>_<index> (default test: Alan_2)
 *  returns {Array of key pairs} selectPerfCharts - perf. charts of those with runner, otherwise all are options 
 */
function GetMyTrendsCharts(
  runnerNameId)
  // runnerNameId ='Alan_13')   // for testing
{
  let runnerTrendCharts = {};
  var trendsTable = [["Trends","=Overview!B$3","","",""]];    // SS Id to span 2 cells propagated into table
  trendsTable.push(["Runner Id","SS Id","Results Gid","#","Shared"]);
  const runnersRANGE = lc.runnersNameCOLUMN+lc.runnersStartROW+":"+lc.runnersNameCOLUMN;
  let runnersNames = lv.allRunnersSheet.getRange(runnersRANGE)
    .getValues()
    .filter(name => name[0] != "")
    .map(name => [name[0]]);
  runnersNames.forEach((runnerName,index) => {
    let runnerId = runnerName+"_"+index;
    let runnerResultsSheet = lv.activeSpreadsheet.getSheetByName(runnerId);
    if (runnerResultsSheet) {
      let runnerShared = getRole(runnerId);
      let runnerNumRuns = getNumRuns(runnerId);
      let runnerResultsSheetId = runnerResultsSheet.getSheetId();
      let runnerResultsCharts = runnerResultsSheet.getCharts();
      // runnerResultsCharts.forEach(trendChart => {
      for (let trendChart of runnerResultsCharts) {

        let chartId = trendChart.getChartId();    // only one, but potentially more
        let chartTitle = trendChart.getOptions().get('title');
        let resultsRange = trendChart.getRanges()[0];
        let chartRange = GetChartRange(resultsRange,1,lc.trendColOFFSET,19,2) // approx 20 rows in height
          .getA1Notation();
        if (runnerNameId)    // filter by runner
          runnerTrendCharts[chartId] = {
            title: chartTitle,
            sheet: runnerId,
            sheetId: runnerResultsSheetId,
            range: chartRange};
        else {  // file all for import
          trendsTable.push([runnerId,"=B$1",runnerResultsSheetId,runnerNumRuns,runnerShared]);
        }
        break;  // first will suffice; others assumed aligned below
      }
    }
  });
  if (runnerNameId) {
    // if (debug)
      Logger.log('Selective Trend charts: for runner, '+runnerNameId+':\n'
        +JSON.stringify(runnerTrendCharts));
    return runnerTrendCharts;
  } else {
    let csv = trendsTable.map(row => row.join(',')).join('\n');
    Logger.log(csv);
    trendsSheet = lv.activeSpreadsheet.getSheetByName('Trends');
    // trendsSheet.clear();  // if required, may already exist in template and pre-formatted; otherwise dynamic
    // let titleRange = trendsSheet.getRange(1,1,1,5);    // table header (below title) has 5 columns
    trendsSheet.getRange(1,1,trendsTable.length,trendsTable[0].length)
      .setValues(trendsTable);
  }
}

/**
 * This prepares the overview (for rank tables), individual trends charts (for each member)
 * with access status), and the performance charts (for filtering links) on a weekly basis
 * after completed import AND after successful generation of all of the comparative charts.
 */
function PrepareAppSheets() {
  GetOverview();        // incl. count of members, parkrun group id & domain
  GetMyTrendsCharts();  // incl. number of runs per member and SS access status
  GetPerformanceCharts(); // 4 divisional categories, e.g. 7 Leagues, 8 Age Groups
}

// Surplus extras may be useful

function PasteAboveRangeFormula() {
  var sheet = lv.activeSpreadsheet.getActiveSheet();
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
  var sheet = lv.activeSpreadsheet.getActiveSheet();
  var activeRange = sheet.getActiveRange();
  var row = activeRange.getRow();
  sheet.insertRowBefore(row);
  var aboveRow = sheet.getRange(row,1,1,sheet.getLastColumn());
  var originalRow = sheet.getRange(row+1,1,1,sheet.getLastColumn());
  originalRow.copyTo(aboveRow,SpreadsheetApp.CopyPasteType.PASTE_FORMULA);
}

function swapColumns() {
  // Swap columns selected by user (potentially across multiple sheets)
  var selection = lv.activeSpreadsheet.getActiveRangeList().getRanges();
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
