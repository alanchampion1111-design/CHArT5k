/* -------------------------------------------------------------------------------------
/
/ This is the client-end GoogleApp Script that resides within the Family Template
/ Google Spreadsheet.  It complements the server side JavaScript index.js script sinc
/ dependent on those asynchronous GCR functions.
/ It primarily consist of a series of five macro-based functions aimed at automating
/ the capture of results (Phase II) including positions (Phase III), and subsequently
/ automating the retro catch-up (Phase IV) and the automation of adding new members
/ (Phase V) as a pre0requisite for start-up of new families/clubs:
/   1.  Import result for each runner     (Phase II & III positions)
/       (ImportResultForEachRunner      - Ctrl+Alt+Shift+0) - weekly
/   2.  Catch-up all positions            (Phase IV)
/       (CatchUpAllPositions            - Ctrl+Alt+Shift+5)
/   3.  Batch positions for runner        (Phase IV)    
/       (BatchPositionsForRunner        - threading)
/   4.  Add family (or club) member       (Phase V)
/       (AddFamilyMember                - Ctrl+Alt+Shift+7)    
/   5.  Spawn new family (or club)        (Phase V)
/       (SpawnNewFamily                 - Ctrl+Alt+Shift+9)
/   6.  Add first member                  (Phase V)
/       (AddFirstMember                 - via library call)
/
/---------------------------------------------------------------------------------------
 */

/**
 * @OnlyCurrentDoc
 *  // Ensure authorisation granted via appsscript.json
 * @scope https://www.googleapis.com/auth/script.external_request
 * @scope https://www.googleapis.com/auth/script.scriptapp
 * @scope https://www.googleapis.com/auth/spreadsheets
 * @scope https://www.googleapis.com/auth/script.container.ui
 * @scope https://www.googleapis.com/auth/drive.readonly
 * @scope https://www.googleapis.com/auth/drive
 */

/* --------------------------------------------------------------------------
/
/   The following definitions and functions automate the addition of a new
/   result by securely connecting to the parkrun site for importing the
/   latest result, mimicking a manual extract to extract copy and paste.
/
/   There are two use-cases, where the heirarchy of function calls are:
/
/   1.  ImportResultsOnEventDate
/         ...to get event Date (dd/MM/yyyy).
/
/   1a. ImportResultForEachRunner (optional event date of Parkrun)
/         GetLastSaturday (to derive last event Date if unspecified)
/         TrackImportDate
/         OpenChromeBrowser ->
/         >GetActiveMember ->
/           >AccessPage
//        Recursively iterate for each active member...
/           >ImportRunnerResult
/           >CopyResultForRunner (dated)->
/             >GetRunnerResultsPage
/               >AccessPage
/             GetResultRow
//          When latest is a new result...
/             >PasteResultForRunner ->
//              Loop for each cell...
/                 CleanValue (apply links as hyperlinks)
/               FormatDate (potentially redundant)
/               AppendResultRow
/             >AppendPositionsForResult ->
/               GetResultUrl (possibly same event URL as previous runner?)
/               GetCategories (for language-based filter)
/               ExtendRange
/               >AssessPositions (for one runner, potentially same page)
/               IncludePositions
/         >CloseChromeBrowser
/
/   2.  CatchUpAllPositions
/         OpenChromeBrowser ->
//        Loop for each member runner...
/           LockCallerForwardsTo ->
/         >CloseChromeBrowser
/
/   2a. BatchPositionsForRunner (threaded in series)
/         UnlockCallerForwarded
/         OpenChromeBrowser ->
/           SyncPositionsPerRunner -> (in batches of 10 results)
/             FirstMatchRange (based on unknown Gender position)
//            When positions not previously updated...
//              Loop for each runner's result (per batch)...
/                 AppendPositionsForResult ->
/                   GetResultUrl
//                  When event is a Parkrun...
/                     GetCategories (for language-based filter)
/                     >AssessPositions (for one runner)
/                     IncludePositions
/                       Trigger forked task, BatchPositionsForRunner
/         >CloseChromeBrowser
/
/   2b. BatchPositionsForRunner (recurse if more to sync):
/         UnlockCallerForwarded
/         When all positions applied...
/           >CleanBatch
/  -------------------------------------------------------------------------
*/

// Global constants and varaibles must be defined within a this file IF potentially a triggered
// TODO: Sync gcrFunctions.gs with localFunctions.gs
const gc = {
  debug: false,               // WARNING: may slow down performance if true
  runnerNameCOLUMN: "A",      // Runners name in column A 
  runnerSurnameCOLUMN: "B",   // Runners surname in column B
  dobINDEX: 4,                // Runners sheet in column E (for arrays or range offsets)
  parkrunnerIdINDEX: 9,       // Runners sheet in column J (for arrays or range offsets)   
  parkrunnerIdCOLUMN: "J",    // Runners parkrunner ID in column J
  hasResultsCOLUMN: "K",      // ...results exist (D3:D), with Parkrunner Id in col J
  hasPosnsCOLUMN: "L",        // ...has Positions up-to-date (I3:I) based on genderPosnCOL
  resultsStartROW: 3,
  runnerNameCELL: 'A1',
  runnersStartROW: 3,
  parkrunnerIdCOL: 10,        // in column J on Runners sheet
  titleNameCELL: "A1",        // Runners cell club/family name (with link to consolidated results)
  templateNameCELL: "E1",     // Runners cell identifies the seed template (e.g. Joe_90)
  importDateCELL: "I1",       // Runners cell for the last import date (d-MMM-yy)
  clubIdCELL: "J1",           // Runners cell identifies the club or family group id
  importIndexCELL: "K1",      // Runners cell with index of runner to continue import
  importTotalCELL: "L1",  // Runners cell with number of runners who ran on this import date
  locationINDEX: 0,       // column A
  dateINDEX: 1,           // column B
  eventCOL: 1,            // column A is the event location
  dateCOL: 2,             // column B is the date of event
  runNumCOL: 3,           // column C is the instance # at event
  PBtickBoxCOL: 8,        // column H tick-box on Results sheets, ticked if value in G is PB
  PBtoAgeCatNumCOLS: 5,   // cols H..L for I..K positions between derived cols, H & L
  genderPosnCOL: 9,       // column I on each Results sheet (until on parkrun results page?)
  ageCatPosnCOL: 10,      // column J on each Results sheet
  ageGradePosnCOL: 11,    // column K on each Results sheet
  ageCatCOL: 12,          // column L is the Age Category on event date (derived from DoB)
  dateFORMAT: 'd-MMM-yy',
  parkrunDateFORMAT: 'dd/MM/yyyy',  // perhaps for UK-based runners?
  universalDateFORMAT: 'yyyy-MM-dd',
  runnersSheetNAME: 'Runners',
  browserURL: 'https://browser-automation-service-224251628103.europe-west1.run.app',    // shared GCR service
  sampleURL: 'https://www.parkrun.org.uk/colchestercastle/results/116',
  batchSizeMAX: 7,
  importTimeSECS: 30,
  maxTimeSECS: 6*60
};
var gv = {
  activeSpreadsheet: SpreadsheetApp
    .getActiveSpreadsheet(),      // allows dynamic shift to new context
  browserSession: undefined,      // Shared by threaded/recursed processes since lingers
  startTime: Date.now(),
  numImports: 0
};
gv.activeSpreadsheetId = gv.activeSpreadsheet.getId();       // the originator ID
gv.allRunnersSheet = gv.activeSpreadsheet   // MUST redo after a dynamic shift during spawning
  .getSheetByName(gc.runnersSheetNAME);     // ...for new runners from initiating Spreadsheet
if (gc.debug) Logger.log('Current Spreadsheet Id: '+gv.activeSpreadsheetId);  

async function OpenChromeBrowser() {
  const initBrowserURL = gc.browserURL+'/initBrowser';
  try {
    var response = await UrlFetchApp.fetch(initBrowserURL);
    browserSession = response.getContentText();   // WS Endpoint lingers for upto 30 minutes
    Logger.log('Session: '+browserSession);
    return true;
  } catch (err) {
    Logger.log(err);
    return false;
  }
}

async function AccessPage(
  thisUrl = gc.sampleURL,
  timeSecs = 20)
{
  const getBrowserURL = gc.browserURL+'/getUrl?url='+thisUrl;
  let timeMax = timeSecs*1000;
  try {
    var response = await UrlFetchApp.fetch(getBrowserURL,
      {muteHttpExceptions:true,timeout:timeMax});
    return response.getContentText();
  } catch (err) {
    Logger.log(err+' within '+thisUrl);
    return null;
  }
}

async function CloseChromeBrowser() {
  const stopBrowserURL = gc.browserURL+'/stopBrowser';
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
const allForONE = '/5k/';
const oneForALL = '/results/';    // TODO: Potentially cache if same location as next member

async function GetRunnerResultsPage(
  parkrunnerId = '1213963', // default useful for testing
  thisPage = undefined)     // already retrieved (short-circuit)
{
  if (thisPage) return thisPage;
  const thisParkrunnerURL = parkrunnerURL+parkrunnerId+allForONE;
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
 * Gets a specific result row for a runner (default to latest in the first row)
 *  @param {Array<string>} bodyRows - Array of HTML rows containing result data.
 *  @param {number|string} [eventRef=0] - Index of the result (0 = latest) or date (dd/MM/yyyy) of result.
 * @returns {string|null} The matching result row HTML, or null if a matching date is not found
 * Assumes operates within "universal" (non-US) date browser setting.
 */
function GetResultRow(
  bodyRows,
  eventRef = 0)  // default to latest (indexed 0) in first row; otherwise match as a date 
{
  if (typeof eventRef === 'number') {
    return bodyRows[eventRef];
  } else {
    return bodyRows.find(row => row.includes(`>${eventRef}<`))
  }
}

/**
 * Copies the latest result for a given parkrunner ID from Parkrun site
 *    @param {string} parkrunnerId - The ID of the parkrunner
 *    @param {string} runnerNameId - skip copy if result on runner's results sheet
 *    @eventRef (string) - ALL or date (dd/MM/yyyy assumes UK format on parkrun site)
 *    @thisPage (html) - typically preloaded when ALL only
 *  @returns {array|null} - represents the latest result (or null if none)
 */
async function CopyResultForRunner(
  parkrunnerId = "777764",    // defaults for testing
  runnerNameId = 'Alan_13',
  eventRef = undefined,   // typically a date (dd/MM/yyyy) otherwise latest
  thisPage = undefined)   // pre-loaded Results Page 
{
  // TODO: 5k result table does not (yet) show Gender position,
  //       ...nor positions for Age-Category nor Age-Grade (%age)
  if (eventRef && eventRef !== 'ALL') {
    // Skip result if already imported as latest, although positions may be unknown
    var resultsSheet = gv.activeSpreadsheet.getSheetByName(runnerNameId);
    // Location & date (cols A..B) should suffice for match & skip PasteResult...
    let latestResult = resultsSheet.getRange(resultsSheet.getLastRow(),1,1,gc.dateCOL)
      .getDisplayValues()[0];   // Assume dated event (if exists) is at the end
    if (latestResult[gc.dateINDEX] === FormatDate(eventRef,gc.dateFORMAT)) {
      if (gc.debug)
        Logger.log('Skip re-importing duplicate (although positions may need updating) for runner, '+runnerNameId);
      return latestResult;
    }
  }
  return GetRunnerResultsPage(parkrunnerId,thisPage)  //  skips if thisPage pre-loaded!!
  .then (allResults => {   // ensures allResults is not a Promise
    if (allResults) {
      let tables = allResults.match(/<table[\s\S]*?<\/table>/ig);
      if (tables && gc.debug) Logger.log('Found ['+tables.length+'] tables in results for runner, '+parkrunnerId);
      if (tables && tables.length>2) {
        allResults = tables[2]; // 5k results (if any) are in 3rd table
        // if (gc.debug) Logger.log('5k Results table: '+allResults); 
        // var headerRow = allResults.match(/<thead>.*?<\/thead>/s);
        // TODO: columns may change - WARN if so?
        var bodyContent = allResults.match(/<tbody>.*?<\/tbody>/s)[0];
        var bodyRows = bodyContent.match(/<tr[^>]*>.*?<\/tr>/gs);
        bodyRows = bodyRows.filter(row => !row.includes('junior'));    // junior 2km times upset rankings
        if (bodyRows && bodyRows.length>0) {
          if (eventRef === 'ALL') {   // return 2D array of results
            return bodyRows.map(row => {
              var cells = row.match(/<td>.*?<\/td>/gs);
              return cells.map(cell => cell.replace(/<td>|<\/td>/g, "").trim());
            });
          } else { // single result, default to latest unless specific date
            var resultRow = GetResultRow(bodyRows,eventRef);  
            if (resultRow) {
              var cells = resultRow.match(/<td>.*?<\/td>/gs);
              if (cells) {
                var values = cells.map(function(cell) {
                  return cell.replace(/<td>|<\/td>/g,"").trim();
                });
                return values;
              } else {
                Logger.log('WARNING: No cells found in row, '+resultRow+' for runner, '+parkrunnerId);
                return null;
              }
            } else if (!eventRef) {
              Logger.log('Unable to find results for runner, '+parkrunnerId);
              return null;
            }
          }            
        } else {
          Logger.log('WARNING: Unable to find results row for runner, '+parkrunnerId);
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
 *    @param {string} dateSource - Date string to format (dd/MM/yyyy expected) 
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

function PrepareResultRow(thisResult) { // cols A..G (7)
  const linksNUM = hyperLinks.length; // cols A..C (3)
  const valsNUM = thisResult.length - linksNUM; // cols D..G (4)
  try {
    thisResult = thisResult.map(cellValue => {
      if (hyperLinks.length) {        // initially, assume 3?
        return '=HYPERLINK("'+
          hyperLinks.shift()          // ..with consuming side-effect
          +'","'+cellValue+'")';
      } else {
        return cellValue;
      }
    });
    let formulae = thisResult.slice(0,linksNUM);  // A-C
    let values = thisResult.slice(linksNUM);    ``// D-G
    return [formulae,values];
  } catch (err) {
    Logger.log('ERROR: Unable to add hyperlinks correctly to latest result for runner...\n'+err);
    return null;
  }
}

/**
 *  Appends a result row (columns A-G) into results sheet (including 2 or more hyperlinks)
 *    @param {string[]} thisResult - Latest result (cell values) for a runner
 *    @param {Sheet} resultsSheet - Sheet to append the latest result row
 *    @param {offset} default append (or update last if first row)
 *  @returns {range} resultsRow - contains the results link to determine detailed positions
*/
// TODO: Replace with PrepareResultRow, and then setValue in the caller (for each runner)
function AppendResultRow(
  thisResult,   // cols A..G (7)
  resultsSheet,
  numRows = 1)
{
  try {
    const linksNUM = hyperLinks.length;             // cols A..C (3)
    const valsNUM = thisResult.length-linksNUM;     // cols D..G (4)
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
    if (numRows === 1)
      if (insertRow > gc.resultsStartROW)
        resultsSheet.getRange(insertRow,gc.PBtickBoxCOL)   // column beyond paste
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
 * Pastes the latest or dated result for a given runner into their sheet.
 *    @param {array} thisResult - representing the latest (single) result
 *    @param {string} [runnerNameId="Alan_13"] - name of the runner's results sheet
 *  @returns {Range} - where result was pasted or had been except positions (otherwise null)
 */
 function PasteResultForRunner(
  thisResult,     // assumes a single row to be added
  runnerNameId = "Alan_13")
{
  if (gc.debug) Logger.log('Consider pasting result into unique runner sheet, '+runnerNameId+'...');
  const previousROWS = 5;   // previous results (for matching) may include non-parkrunner events
  hyperLinks = [];   // discard any hyperLinks from duplicate results
  thisResult = thisResult.map(CleanValue);   // includes 00: (for hh:) prefix on elapsed time
  var thisDate = thisResult[gc.dateINDEX] = FormatDate(thisResult[gc.dateINDEX],gc.dateFORMAT);
  var thisLocation = thisResult[gc.locationINDEX];
  var resultsSheet = gv.activeSpreadsheet.getSheetByName(runnerNameId);
  var maxRows = Math.min(previousROWS,resultsSheet.getLastRow()-gc.resultsStartROW+1); // e.g. 1 if only 3 rows
  var firstPrevRow = resultsSheet.getLastRow()-maxRows+1;   // e.g. 18=22-5+1 (from 19th row if last is 23)
  var previousResultsRange = resultsSheet.getRange(firstPrevRow,1,maxRows,resultsSheet.getLastColumn());
  var previousResults = previousResultsRange.getDisplayValues();
  // if (gc.debug) Logger.log('Match previous results:\n'+previousResults);
  if (previousResults.length < 1) {
    Logger.log('ERROR: No previous result for unique runner, '+runnerNameId+' to be able to match new result');
    return null;
  }
  let matchingResults = previousResults.map(function(row) {
    return row[gc.locationINDEX]+'&'+row[gc.dateINDEX];
  });
  let matchIndex = matchingResults.indexOf(thisLocation+'&'+thisDate);  // check duplicate?
  if (matchIndex > -1) {   // existing result, but is it complete?
    hyperLinks = [];  // discard pending hyperlinks since redundant for duplicate result
    if (gc.debug)
      Logger.log('Previously captured result at '+thisLocation+' on '+thisDate+
        ' for unique runner, '+runnerNameId);
    let partialResult = previousResultsRange.offset(matchIndex,0,1);
    // Check whether existing result has had positions
    let genderPositionKnown = previousResults[matchIndex][genderPosnCOL-1];  // 0-indexed
    // let genderPositionKnown = partialResult.getCell(1,genderPosnCOL).getValue();
    if (gc.debug)
      Logger.log('genderPositionKnown [row]: '+genderPositionKnown
        +' ['+matchIndex+' from '+firstPrevRow+']');
    if (genderPositionKnown !== "") {
      Logger.log('Previously imported result for '+thisLocation+' on '+thisDate+
        ' already added (with positions) for unique runner, '+runnerNameId);
      return null;    // No update needed if positions already known
    } else {
      if (gc.debug)
        Logger.log('Matched result at '+thisLocation+' on '+thisDate+
          ' for unique runner, '+runnerNameId+' except for positions');
      return partialResult;   // continue as if new result because positions omitted
    }
  } else {    // new result
    var pastedResult = AppendResultRow(thisResult,resultsSheet);
    Logger.log('New result imported for '+thisLocation+' on '+thisDate+
      ' added for unique runner, '+runnerNameId);
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
async function AssessPositions(matchRunner,thisUrl,ageCat,genderCat,cacheUrl=false) {
  const filterBrowserURL = gc.browserURL+'/filterUrl'
    +'?url='+thisUrl
    +'&rn='+encodeURIComponent(matchRunner)
    +'&ac='+ageCat
    +'&gc='+genderCat
    // +'&ag='Age-Grade'   // assume internal default
    +'&cache='+cacheUrl;   // only cache on import for event-date / latest date
  try {
    var positions = await UrlFetchApp.fetch(filterBrowserURL,
      {muteHttpExceptions:true,timeout:25000});
    let posnText = positions.getContentText();
    if (gc.debug) Logger.log('Runner: '+matchRunner+'\t'+posnText);
    // TODO Perhaps redundant with managable batch size (<20) unless browser close prematurely!!
    /* if (posnText.includes('Internal Server Error')) {
      browserSession = undefined; // restart the browser!
      Logger.log('CAUTION: Check GCR logs for unexpected errors');
      return CloseChromeBrowser()
      .then(() => {
        OpenChromeBrowser()
      });
      return null;
    }
    */
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
  resultRange.offset(0,genderPosnCOL-1,1,3).setValues([[gcPosn,acPosn,agPosn]]);
}

/**
 * Extends the runner's result range passively, typically one result in a single row
 *  @param {Range} resultRange 
 *  @param {numeric} numRows 
 *  @param {string} agPosn
 */
function ExtendRange(resultRange,numRows,numCols) {
  // Two derived values are generated from a MAP function set for the first result in the table 
  resultRange = resultRange.offset(0,0,resultRange.getNumRows()+numRows,resultRange.getNumColumns()+numCols);
  // resultRange.getCell(resultRange.getNumRows(),ageCatCOL)
  //  .setValue(null);  // no conflict in derivation from event date in col B and runner's DoB
  if (gc.debug) Logger.log('Extended range of results by '+numRows+' rows and '+numCols+' columns');
  // resultRange implicitly includes  derived Age-Category in col L (assumes MAP function in 1st result)
  let ageCatCell = resultRange.getCell(1,ageCatCOL);
  let ageCatCellRef = ageCatCell.getA1Notation();
  let ageCat = ageCatCell.getValue();
  if (gc.debug) Logger.log('Age-Category is '+ageCat+' in new runner row at cell, '+ageCatCellRef);
  return resultRange;
}

/**
 * Gets the URL for the event from the hyperlink (against the date in Col B)
 * Otherwise, legacy solution uses event data and instance from neighbouring cells 
 * Handles specific non-UK event locations by mapping them to their respective parkrun domains.
 * Handles anomalies on some Parkrun Parks including/excluding 'Park' in parkrun domains
 */
function GetResultsUrl(eventDateCell,eventLocation,eventInstance) {
  var resultsLink;
  const eventDate = eventDateCell.getDisplayValue;
  eventDateCell.activate();   // hyperlink may surface when cell selected?
  let richText = eventDateCell.getRichTextValue();
  if (richText) {   // WARNING: Only works for cells with HYPERLINK formulae
    resultsLink = richText.getLinkUrl();
  } else {   // REDUNDANT: Legacy complication for past results, the embedded URL was not retrievable!!
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
  }
  resultsLink = (resultsLink.includes('parkrun')) ? resultsLink : null;
  return resultsLink;
}

// Since 2026, these three MAP definitions are now needed to correspond to the correct gender and age-category
// Assume derived from age category: e.g. SW25-29 for native UK => SV25-29 in NL event (and vice-versa)
// TODO: For international, better to inspect the gender position directly from detailed result?
const genderMAP = {
  'en': { M: 'Male', W: 'Female' },       // confirmed
  'nl': { M: 'Man', V: 'Vrouw' },         // confirmed
  'de': { M: 'Männlich', F: 'Weiblich' }, // confirmed
  'it': { M: 'Uomo', W: 'Donna' },        // confirmed
  'fr': { H: 'Homme', F: 'Femme' },
  'da': { M: 'Mand', W: 'Kvinde' },
  'fi': { M: 'Miehet', N: 'Naiset' },     // confirmed
  'lt': { V: 'Vyras', M: 'Moteris' },     // confirmed
  'no': { M: 'Mann', K: 'Kvinne' },       // confirmed
  'pl': { M: 'Mężczyzna', W: 'Kobieta' },
  'se': { M: 'Man', K: 'Kvinna' },
  'jp': { M: '男子', W: '女子' }    // pronounced Danshi or Joshi
};

const toUK= {
  'en': { 'M': 'M', 'W': 'W' },     // confirmed
  'nl': { 'M': 'M', 'V': 'W' },     // confirmed
  'de': { 'M': 'M', 'F': 'W' },     // confirmed
  'it': { 'M': 'M', 'W': 'W' },
  'fr': { 'H': 'M', 'F': 'W' },
  'da': { 'M': 'M', 'W': 'W' },     // confirmed as en
  'fi': { 'M': 'M', 'N': 'W' },     // confirmed
  'lt': { 'V': 'M', 'M': 'W' },     // confirmed
  'no': { 'M': 'M', 'K': 'W' },
  'pl': { 'M': 'M', 'W': 'W' },
  'se': { 'M': 'M', 'K': 'W' },
  'jp': { 'M': 'M', 'W': 'W' }      // confirmed as en
};

// Together with the inverse of toUK allows a native gender abbrev. to be translated
// to any other event country gender in two steps (via this 2-way UK mapping)
let fromUK = {};    
Object.keys(toUK).forEach(lang => {
  fromUK[lang] = {};
  Object.entries(toUK[lang]).forEach(([local,uk]) => {
    fromUK[lang][uk] = local;
  });
});

const countryMAP = {
  'uk': 'en',
  'ie': 'en',
  'us': 'en',
  'au': 'en',
  'za': 'en',
  'at': 'de',
  'nl': 'nl',
  'de': 'de',
  'it': 'it',
  'fr': 'fr',
  'dk': 'da',
  'fi': 'fi',
  'no': 'no', // without any stage prefix char (omit J, S and V)
  'pl': 'pl',
  'se': 'se', // without stage prefix char for J(unior) or S(enior)
  'jp': 'jp',
  'ca': 'en', // or 'fr' for French-speaking Canada?
  'sg': 'en',
  'nz': 'en',
  'my': 'en', // or 'ms' for Malay
  'na': 'en',
  'lt': 'lt',
  'es': 'es'    // Add Spain to other MAPs when launched
};

/**
 * Returns the gender and age category based on the URL domain of the event
 * which may differ from the native country of the spreadsheet
 *    @param {string} [eventUrl=parkrunURL] - URL to determine domain
 *    @param {string} [thisAgeCat='VW35-39'] - Gender code (M/U/H or W/V/K/D/F)
 *  @returns {array} [Gender code (e.g. "Male" for M, "Vrouw" for V, etc.), Event-based age-category]
 */
function GetCategories(
  eventUrl = parkrunURL,      
  thisAgeCat = 'VW35-39')   // as on (native) spreadsheet
{
  // TODO: retrieve native Domain/Language from  the spreadsheet (owner home)
  const nativeDOMAIN = 'uk';
  const nativeLANG = countryMAP[nativeDOMAIN];              // i.e. en
  let eventDomain = eventUrl.split('/')[2];
  let eventCountry = eventDomain.split('.').pop();          // e.g. at
  let eventLang = countryMAP[eventCountry] || eventCountry; // e.g. de
  let newAgeCat = thisAgeCat;   // default to no change
  if (nativeLANG !== eventLang) {   // convert native -> UK -> event,
    // ...assumes prefix is V(eteran), S(enior), or J(unior), even in Japan!
    let thisGender = thisAgeCat[1];
    newAgeCat = thisAgeCat.replace(thisGender,          // e.g. W => V
      fromUK[eventLang][toUK[nativeLANG][thisGender]]   // VW35-39 => VV35-39
    );
  }
  let genderWord = genderMAP[eventLang][newAgeCat[1]];
  // strip stage prefix (J, S or V) if Norway, or if Sweden unless Veteran
  if (eventLang === 'no' || (eventLang === 'se' && /^[JS]/.test(newAgeCat)))
    newAgeCat = newAgeCat.slice(1);
  return [genderWord,newAgeCat];
}

/**
 * Appends detailed positions to an individual runner's sheet beyond the current range.
 * Used initially where region needs extended OR later on catch up when already so.
 *    @param {string} runnerFullName - Runner's full name
 *    @param {Range} resultRange - Range of result data
 *    @param {number} runnerIndex - index of runner imported to date (unless -1 as default)
 *    @param {voolean} cacheUrl - cache as separate Page for re-use (REDUNDANT?)
 *  @returns {boolean} Success flag
 */
async function AppendPositionsForResult(runnerFullName,resultRange,
  runnerIndex = -1,   // by default, do NOT track importing positions for all runners
  cacheUrl = false)
{
  const runNumCOL = 3;          // column C is the instance # at event
  let runNumber = resultRange.getCell(1,runNumCOL).getValue();
  if (!runNumber) return false;   // Skipping any non-Parkrun instance since no positions
  else {
    let eventLocation = resultRange.getCell(1,gc.eventCOL).getValue();
    let eventDateCell = resultRange.getCell(1,gc.dateCOL);
    let extResultRange = (resultRange.getNumColumns() < gc.ageCatCOL) 
      ? ExtendRange(resultRange,0,gc.PBtoAgeCatNumCOLS) // extends to include H..L (for Import use-case)
      : resultRange;             // Result row range previously extended (for Catch-up use-case)
    let genderPositionKnown = extResultRange.getCell(1,genderPosnCOL).getValue();  
    if (genderPositionKnown) {
      if (runnerIndex >= 0)
        gv.allRunnersSheet.getRange(gc.importIndexCELL).setValue(runnerIndex+1);
      return false;  // Skipping since extra position(s) already on the sheet
    } else {
      try {
        let ageCategory = extResultRange.getCell(1,ageCatCOL).getValue();  // derived from Date - DoB
        let genderCategory;  // = runnersSHEET.getCell(runnerIndex,genderCatCOL).getValue();  // out of Range!!
        var resultsLink = GetResultsUrl(eventDateCell,eventLocation,runNumber); // TODO: since workaround not needed
        // var resultsLink = GetResultsUrl(eventDateCell); // TODO: behind Date (in col B)
        // Adjust ageCategory to suit language at event - e.g. SW25-29 (uk) => SV25-29 (de)
        [genderCategory,ageCategory] = GetCategories(resultsLink,ageCategory); // e.g. JM10 => M => Male in uk
        if (gc.debug)
          Logger.log('Link to '+eventLocation+' Parkrun results:\n'+resultsLink);
        let extraPosns = await AssessPositions(runnerFullName,resultsLink,ageCategory,genderCategory,cacheUrl);
        if (extraPosns && extraPosns.length === 3) {
          Logger.log('Positions captured for '+runnerFullName+' ('+resultsLink+'): '+extraPosns);
          IncludePositions(extResultRange,...extraPosns);  // into cells, I..K on result row
          if (runnerIndex >= 0) gv.allRunnersSheet.getRange(gc.importIndexCELL).setValue(runnerIndex+1);
          return true;
        } else {
          Logger.log('CAUTION: Consider revising DoB of '+runnerFullName+' to ensure matching category ('+ageCategory+') at '+eventLocation+' ('+resultsLink+')?');
          return false;
        }
      } catch (err) {
        Logger.log('ERROR: While appending positions for runner, '+runnerFullName+': '+err);
        return false; // or throw error
      }
    }
  }
}

function GetLastSaturday(date) {
  // var d = new Date();
  date.setDate(date.getDate()-(date.getDay()+1)%7); // last Saturday
  return Utilities.formatDate(date,Session.getScriptTimeZone(),gc.parkrunDateFORMAT)
}

function TrackImportDate(eventDate) {
  let startIndex = 0;
  let prevImportDate = gv.allRunnersSheet.getRange(gc.importDateCELL).getDisplayValue();
  if (gc.debug) Logger.log('Import date: '+eventDate+' ('+prevImportDate+')');
  if (eventDate === prevImportDate) {   // continue since assume prev import incomplete
    startIndex = gv.allRunnersSheet.getRange(gc.importIndexCELL).getValue();
  } else {  // fresh import
    gv.allRunnersSheet.getRange(gc.importDateCELL).setValue(eventDate);
    gv.allRunnersSheet.getRange(gc.importIndexCELL).setValue(startIndex);
  }
  if (startIndex == 0)  // whether a fresh import or manually reset
    gv.allRunnersSheet.getRange(gc.importTotalCELL).setValue(0);
  return startIndex;
}

async function GetEventsResults(eventDate) {
  let effectiveDate = FormatDate(
    eventDate.toString(),   // dd/MM/yyyy not recognised as a date constructor
    gc.universalDateFORMAT
  );
  let clubNameCell = gv.allRunnersSheet.getRange(gc.titleNameCELL);
  // var clubName = clubNameCell.getValue()  // // e.g. Overton Harriers & AC
  //  .replace(clubSUFFIX,"").trim().replace('&','&amp;');
  let clubWideResultsUrl = clubNameCell.getRichTextValue()
    .getLinkUrl();   // hyperlink to consolidated results in A1
  var eventsResults;
  if (clubWideResultsUrl) { // most recent results without a date
    let effectiveUrl = clubWideResultsUrl+'&eventdate='+effectiveDate;
    let clubWideResults = await AccessPage(effectiveUrl,30);
    clubWideResults = clubWideResults
      .slice(clubWideResults.indexOf('<div class="floatleft">'),
        clubWideResults.indexOf('<div class="results-error">'));
    eventsResults = clubWideResults.match(/<table[^>]*>(.*?)<\/table>/gs);
    if (gc.debug)
      Logger.log('No. of events covered: '+eventsResults.length);
  } else
    Logger.log('WARNING: No defined group OR consolidated report for date does not exist;'+
      ' thus checking all runners individually');
  return eventsResults;
}

async function ImportRunnerResult(index,runners,eventDate,eventsResults) {
  let runner = runners[index];
  let runnerName = runner[0];     // col A for first name
  let runnerNameId = runnerName+'_'+index;
  let runnerFullName = runner.join(' ');    // from col A & B
  let parkrunnerId = gv.allRunnersSheet
    .getRange(gc.runnersStartROW+index,gc.parkrunnerIdCOL)
    .getValue();
  if (isNaN(parkrunnerId)) {
    if (gc.debug) Logger.log('WARNING: Skipping invalid parkrunner Id, '+parkrunnerId);
    return true;    // skip
  } else if (eventsResults &&
            !eventsResults.some(table => table.includes(runnerFullName))) {
    if (gc.debug)
      Logger.log('INFO: Skipping runner, '+runnerFullName+' because did not run or not a member');
    return true;    // skip
  } else
    if (gc.debug)
      Logger.log('Parkrunner ID: '+parkrunnerId);
  return CopyResultForRunner(parkrunnerId,runnerNameId,eventDate)
    .then(thisResult => {
      if (thisResult) {
        let resultRange = PasteResultForRunner(thisResult,runnerNameId);
        if (resultRange) {
          gv.numImports++;
          if (gc.debug)
            Logger.log('Appending result positions for unique runner sheet, '+runnerNameId);
          return AppendPositionsForResult(runnerFullName,resultRange,index,true); // caching
        } else {
          gv.allRunnersSheet.getRange(importIndexCELL).setValue(index);
          if (gc.debug)
            Logger.log('Positions already appended for runner, '+runnerNameId);
          return true;
        }
      } else {    // provided no timeout or other error...
        if (gc.debug)
          Logger.log('Runner, '+runnerName+' likely did not run on '+eventDate);
        return false;   // and therefore does not increment the numWhoRan
      }
    })
    .then((success) => {
      if (success) {
        let numWhoRan = gv.allRunnersSheet.getRange(gc.importTotalCELL).getValue();
        gv.allRunnersSheet.getRange(gc.importTotalCELL).setValue(numWhoRan+1);
      }
      return true;
    });
}

/**
 *  Imports the latest results for each of member runners from Parkrun site,
 *  potentially on a specific date (if scheduled or missed), otherwise the latest is assume
 */
function ImportResultForEachRunner(
  // potentially import missing results for a date = e.g. '27/12/2025' or '01/01/2026'
  eventDate = undefined)  // undefined means latest Saturday - return to this state otherwise
{
  var date;
  if (typeof eventDate === 'object' && eventDate !== null && 'authMode' in eventDate)   // activated by trigger
    date = new Date(
      eventDate.year,
      eventDate.month-1,
      eventDate['day-of-month'],
      eventDate.hour,eventDate.minute,eventDate.second
    );
  else if (!eventDate)
    date = new Date();
  if (date)
    eventDate = GetLastSaturday(date);
  var startIndex = TrackImportDate(eventDate);
  if (gc.debug)
    Logger.log('Checking for results on event date, '+eventDate
      +' (from runner index, '+startIndex+')');
  return OpenChromeBrowser()
  .then(() => {   // Browser always launched beforehand...
    let runners = gv.allRunnersSheet.getRange(
      gc.runnerNameCOLUMN+gc.runnersStartROW+":"+gc.runnerSurnameCOLUMN
    ).getValues().filter(String);
    return GetEventsResults(eventDate)
      .then(eventsResults => { 
        // EITHER start afresh (new date, index 0) OR continue from previous import (same datem next index)
        let promise = Promise.resolve();  // a promise avoids need for recursion to enforce sequential import
        for (let index=startIndex; index<runners.length; index++) {
          promise = promise.then(() =>
            ImportRunnerResult(index,runners,eventDate,eventsResults))
              .then(() => {
                if (index+1 == runners.length)
                  return; // skip considering need for further import if already on the last runner
                let remainingTimeSecs = gc.maxTimeSECS-Math.floor((new Date().getTime()-gv.startTime)/1000);
                if (gc.debug)
                  Logger.log('Script time remaining: '+remainingTimeSecs+' seconds');
                if (remainingTimeSecs < gc.importTimeSECS) {
                  Logger.log('WARNING: Insufficient time ('+remainingTimeSecs+' secs) to import more results.');
                  gv.allRunnersSheet.getRange(importIndexCELL).setValue(index + 1); // store next index
                  Logger.log('Exiting loop and continuing import beyond index, '+index);
                  return Promise.reject('Avoid timeout');
                }
              });
        }
        return promise
          .then(() => {
            if (gc.debug)
              Logger.log('Completed number of imports ('+gv.numImports+')')
          })
          .catch(esc => {
            if (esc === 'Avoid timeout') {
              ScriptApp.newTrigger('ImportResultForEachRunner')
                .timeBased()
                .after(10*1000) // resume in 10 seconds
                .create();
            } else {
              throw esc;
            }
          });
      })
  })
  .catch(error => {
    Logger.log('ERROR: '+error);
  })
  .finally(() => {
    let numWhoRan = gv.allRunnersSheet.getRange(gc.importTotalCELL).getValue();
    Logger.log('Number of results imported/updated: '+gv.numImports+
      ' (out of '+numWhoRan+' who ran on '+eventDate+')');
    return CloseChromeBrowser();
  });
}

/**
 * Imports results for each runner on a specific date
 *  1. Prompts for the date
 *  2. Imports result for each runner (on that date)
 */
function ImportResultsOnEventDate() {
  const edRegexp = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
  const edPlace = gc.parkrunDateFORMAT;   // dd/MMM/yyyy on parkrun site
  const ui = SpreadsheetApp.getUi();
  const formTitle = 'Import Results on Event Date';
  const formDesc = 'This imports the result for any member who ran on a particular date.';
  const formAction = 'Import';
  const formHandler = 'ImportResultForEachRunner';
  const formHTML = '\n'+
    '<form onsubmit="if(!document.getElementById(\'eventDate\').checkValidity()) {'+
    '    document.getElementById(\'eventDate\').focus();return false;}">\n'+
    '  <div>'+formDesc+'</div><br>\n'+
    '  <label>Date of parkrun event:\t</label>\n'+
    '    <input type="text" id="eventDate" pattern="'+edRegexp.source+'" placeholder="'+edPlace+'"><br><br>\n'+
    '  <input type="submit" id="submitButton" value="'+formAction+'"\n'+
    '    onclick="document.getElementById(\'submitButton\').disabled=true;\n'+
    '      document.getElementById(\'submitButton\').value=\'Processing...\';\n'+
    '    google.script.run.withSuccessHandler(function() { google.script.host.close(); }).'+
          formHandler+'(document.getElementById(\'eventDate\').value)">\n'+
    '  <input type="button" value="Cancel"\n'+
    '    onclick="google.script.host.close()">\n'+
    '  </form>\n';
  var form = ui.showModalDialog(HtmlService.createHtmlOutput(formHTML),formTitle);
  // return eventDate;  // an array of 1 value?
  // callback with eventDate (string) from form
}

function AppendAllResults(
  runnerNameId,
  linksCount,allFormulae,
  valuesCount,allValues)
{
  var resultsSheet = gv.activeSpreadsheet.getSheetByName(runnerNameId);
  let numResults = allFormulae.length;    // = allValues.length
  resultsSheet.getRange(gc.resultsStartROW,1,numResults,linksCount)
    .setFormulas(allFormulae);
  resultsSheet.getRange(gc.resultsStartROW,linksCount+1,numResults,valuesCount)
    .setValues(allValues);
  resultsSheet.getRange(gc.resultsStartROW+1,PBtickBoxCOL,numResults-1,1)   // column beyond paste
    .clear({contentsOnly: true})  // complement to MAP prevents FALSE 
    .insertCheckboxes();          // c/fwd tickbox restored
}

function PasteAllResultsForRunner(
  runnerNameId = 'Sarah_17',
  allResults)
{
  let linksCount = 3;
  let valuesCount = 4;
  let allFormulae = [];
  let allValues = [];
  // Paste in reverse order from parkrun site => earliest result first
  allResults.reverse().forEach(thisResult => {
    thisResult = thisResult.map(CleanValue);        // extracts hyperLinks...
    thisResult[gc.dateINDEX] = FormatDate(thisResult[gc.dateINDEX],gc.dateFORMAT);
    let [rowFormulae,rowValues] =
      PrepareResultRow(thisResult);   // ...applies hyperLinks
    linksCount = rowFormulae.length;   // constant = # of formulae
    valuesCount = rowValues.length;   // constant = # of non-formulae (4)
    allFormulae.push(rowFormulae);
    allValues.push(rowValues);
  });
  AppendAllResults(runnerNameId,linksCount,allFormulae,valuesCount,allValues);
}

/**
 * Imports all results for a given parkrunner ID, pasting them into the sheet.
 * Note: Does NOT append positions (done later in Step 4 catch-up).
 *    @param {string} parkrunnerId - The ID of the parkrunner
 *  @returns {Promise} Resolves when all results are pasted
 * Assumes browser already opened
 */
async function ImportAllResultsForRunner(
  parkrunnerId = '11306668',
  runnerNameId = 'Jonah_17',
  thisPage = undefined    // expected known when 'ALL' (since required to get name)
) {
  if (gc.debug) Logger.log('Import All from parkrunner, '+parkrunnerId+' to results sheet, '+runnerNameId);
  return CopyResultForRunner(parkrunnerId,runnerNameId,'ALL',thisPage)
    .then(allResults => {
      if (allResults) {
        PasteAllResultsForRunner(runnerNameId,allResults);
      }
      return true;
    })
    .catch(err => 
      Logger.log('ERROR: Importing all results\n'+err));
}

/**
 * Creates a new results sheet for a runner if one doesn't exist.
 *    @param {Array of strings} runnerNames - Two-part name of the runner
 *    @param {String} gender, Male / Female
 *    @param {String} email address
 *    @param {String} DoB
 *    @param {number} parkrunnerId - The Park Runner ID
 * @return {String} the name of the new results sheetm based on first name and row number
 */
function CreateRunnerResultsSheet(
  runnerNames,gender,
  email,dob ,  // otherwise null dob if runnerIndex row exists with these details
  parkrunnerId,
  runnerIndex = undefined   // add  to bottom (index) in Runners sheet unless updating 
)
{
  if (gender && dob) {    // for new members except the first runner entry 
    rangeRow = gv.allRunnersSheet.appendRow(
      [...runnerNames,        // A..B
      gender,email,dob,       // C..E
      undefined,undefined,undefined,undefined,  // F..I since derived categories
      parkrunnerId,null,null] // J..L parkrun + derived tick boxes: has results. has positions
    );
    runnerIndex = gv.allRunnersSheet.getLastRow()-gc.runnersStartROW;
  } else {
    // Assume details already on Runners Sheet for first runner
  }
  // Create runner's results sheet if it doesn't exist
  let [runnerName,runnerSurname] = runnerNames;
  let runnerNameId = runnerName+'_'+runnerIndex;
  let templateResults = gv.activeSpreadsheet.getSheetByName(templateNAME);
  let newResultsSheet = templateResults.copyTo(gv.activeSpreadsheet).setName(runnerNameId);
  // ensure the content of the new sheet is unique
    let runnerFullName = runnerNames.join(" ");
    const allRunnersGID = gv.allRunnersSheet.getSheetId();
    newResultsSheet.getRange(gc.titleNameCELL)
      .setFormula('=HYPERLINK("#gid='+allRunnersGID+'","'+runnerFullName+'")');
  // return to Runners sheet and add the link back and the fast results link
    let newResultsGid = newResultsSheet.getSheetId();
    let parkrunnerResultsUrl = parkrunnerURL+parkrunnerId+gc.allForONE;
    rangeRow = gv.allRunnersSheet.getLastRow();
    gv.allRunnersSheet.getRange(rangeRow,1)     // Col. A
      .setFormula('=HYPERLINK("#gid='+newResultsGid+'","'
        +runnerName+'")');
    gv.allRunnersSheet.getRange(rangeRow,gc.parkrunnerIdCOL)     // Col. J    
      .setFormula('=HYPERLINK("'+parkrunnerResultsUrl+'",'
        +parkrunnerId+')');
    if (runnerIndex != 0) {
      let numCols = gv.allRunnersSheet.getLastColumn();
      gv.allRunnersSheet.getRange(rangeRow-1,1,1,numCols)
        .copyFormatToRange(gv.allRunnersSheet,1,numCols,rangeRow,rangeRow);
    }
  return runnerNameId;    // name of the newly created sheet
}

/*
/  Hierarchy for two use-cases prompting for a runner:
/
/   1.  AddFamilyMember (as a member of your family/club)
/         PromptForRunner(addCASE)-> to get parkrun id, Dob,..
/         :
/       DoAddFamilyMember
/         OpenChromeBrowser->   
/         >GetRunnerResults (to get Name from the Results Page)
/           AccessPage (c/fwd...)
/         GetRunnerDetails
/         CreateRunnerResultsSheet (in current spreadsheet)
/         ImportAllResultsForRunner (...b/fwd page)
/           CopyResultForRunner (ALL results from page)
/           PasteAllResultsForRunner (chronologically in cs)
/             CleanValue
/             FormatDate (TODO: potentially redundant)
/             PrepareResultRow
/             AppendAllResults
/         LockCallerForwardsTo-> (triggers)
//          Trigger task, BatchPositionsForRunner...
/
/   2.  SpawnNewFamily (to be owned by first member)
/         PromptForRunner(spawnCASE)-> to get parkrun id, Dob,..
/         :
/       DoSpawnNewFamily
/         OpenChromeBrowser->
/         >GetRunnerResults (to get Name from the Results Page)
/           AccessPage (c/fwd...)
/         GetRunnerDetails
/         >InstantiateFamilySpreadSheet (new spreasdsheet, fs)
/           CreateNewSpreadsheet
/         >AddFirstMember (as a member of new family/club)
/         CreateRunnerResultsSheet (in fs, with Name & index 0)
/         ImportAllResultsForRunner (in fs, ...b/fwd page)
/           CopyResultForRunner (ALL results from page)
/           PasteAllResultsForRunner (chronologically in fs)
/             CleanValue
/             FormatDate (potentially redundant)
/             PrepareResultRow
/             AppendAllResults
/         LockCallerForwardsTo-> (triggers)
//          Trigger task, BatchPositionsForRunner...
/
/   3.  DeleteFamilyMember (as a member of your family/club)
/         PromptForRunner(deleteCASE)-> to get parkrun id, Dob,..
/         :
/       DoDeleteFamilyMember
/         FindRunnerIndex (matching parkrunner Id)
/         If runner found in Runners table matching DoB on the form...
/           ...delete runner's results page
/           ...delete ruuner's (indexed) row in Runners table
/           ...rename runners (indexed) sheets below that
/
*/

function GetRunnerDetails(thisPage) {
  if (!thisPage)
    throw new Error('ERROR: Delayed or unable to access Runner results');
  const nameREGEXP = /<h2>(.*?)<\/h2>/;
  let runnerFullName = (thisPage.match(nameREGEXP) || [])[1]
    .replace(/<span.*?<\/span>/,'')
    .replace(/&nbsp;/g,'')
    .trim();
  if (gc.debug) Logger.log ('3b. Name: '+runnerFullName);
  let runnerNames = runnerFullName.split(' ');
  if (runnerNames.length > 2) {   // Keep any middle name with surname 
    runnerNames = [runnerNames[0],runnerNames.slice(1).join(' ')];
  }
  const paraREGEXP = /<p>(.*?)<\/p>/s;
  let paraContent = (thisPage.match(paraREGEXP) || [])[1];
  let category = paraContent.trim().split(' ').pop();
  var gender = category
    ? ( category.includes('M') ? 'Male'
      : category.includes('W') ? 'Female'
      : null)
    : null;
  return [runnerNames,gender]; 
}

const addCASE = {
  title: 'Add Family / Club Member',
  desc:  'This adds a new member parkrunner into the current Spreadsheet and captures their results. '
        +'The runner is identified by a new row in the Runners sheet plus a separate sheet for the '
        +'results of that new parkrun family member (based on their first name and unique row index), '
        +'where the detailed positions of those results will eventually appear from a background task.',
  action: 'Add',
  handler: 'DoAddFamilyMember'
};

function PromptForRunner(
  thisCase = addCASE)
{
  const ui = SpreadsheetApp.getUi();
  const dobRegex = /^(0?[1-9]|[12][0-9]|3[01])-[A-Za-z]{3}-(19|20)?\d{2}$/;
  const dobPlace = 'dd-Mmm-yyyy';
  var formHTML = '\n'+
    '<form onsubmit="if(!document.getElementById(\'dob\').checkValidity()) {'+
    '    document.getElementById(\'dob\').focus();return false;}">\n'+
    '  <div>'+thisCase.desc+'</div><br>\n'+
    '  <label>parkrun Id (barcode numeric part only):\t</label>\n'+
    '    <input type="number" id="parkrunnerId"><br><br>\n'+
    '  <label>Date of birth (for verifying age-grade):\t</label>\n'+
    '    <input type="text" id="dob" pattern="'+ dobRegex.source+'" placeholder="'+dobPlace+'"><br><br>\n';
  if (thisCase != deleteCASE) formHTML += 
    '  <label>Email (optional, for delegating access):\t</label>\n'+
    '    <input type="text" id="email"><br><br>\n';
  formHTML +=
    '  <input type="submit" id="submitButton" value="'+thisCase.action+'"\n'+
    '    onclick="document.getElementById(\'submitButton\').disabled=true;\n'+
    '      document.getElementById(\'submitButton\').value=\'Processing...\';\n'+
    '    google.script.run.withSuccessHandler(function() { google.script.host.close(); }).'+
          thisCase.handler+'([document.getElementById(\'parkrunnerId\').value,\n'+
    '      document.getElementById(\'dob\').value';
  if (thisCase != deleteCASE) formHTML +=
    ',\n      document.getElementById(\'email\').value';
  formHTML +=
    ']);">\n         <input type="button" value="Cancel"\n'+
    '    onclick="google.script.host.close();">\n'+
    '  </form>\n';
  var form = ui.showModalDialog(HtmlService.createHtmlOutput(formHTML),thisCase.title);
  // return [parkrunnerId,dob,email];
}

/**
 * Find matching runner index in Runners sheet range
 *    @param {Range} runnersRange - 
 *    @param {number} thisParkrunnerId - parkrunner barcode (numeric!)
 *    @param {Date} thisDob (string) - form format, d-Mmm-yy or dd-Mmm-yyyy
 *  @returns {number} index - unique index (0+) in Runners sheet (-1 if rare no match)
 */
function FindRunnerIndex(runnersRange,thisParkrunnerId,thisDob) {
  let numRrunners = runnersRange.getNumRows();
  let timestamp = new Date(thisDob); timestamp = timestamp.getTime();   // safe, as separate steps
  if (isNaN(thisParkrunnerId)) 
    Logger.log('ERROR: Numeric barcode expected for parkrunner ['+
      typeof thisParkrunnerId+' to ensure matching');
  else for (let index=0; index<numRrunners; index++) {
    let parkrunnerId = runnersRange.offset(index,gc.parkrunnerIdINDEX).getValue();
    let dob = runnersRange.offset(index,gc.dobINDEX)
      .getValue().getTime();    // timstamp
    if (parkrunnerId === thisParkrunnerId)
      if (dob === timestamp)
        return index;
      else
        Logger.log('WARNING: Matched runner ['+parkrunnerId+
          '] expects matching DoB? [not '+thisDob+']');
  }
  return -1;  // no match
}

const deleteCASE = {
  title: 'Delete Family / Club Member',
  desc:  'This deletes an existing member parkrunner from the current Spreadsheet with related impact. '
        +'The runner row is removed from the Runners sheet with their sheet for their results (based '
        +'on row index), where indices and result sheet names of runners below are affected.',
  action: 'Delete',
  handler: 'DoDeleteFamilyMember'
};

/**
 * Deletes an existing runner from the current sheet after prompting for details:
 *  1. Prompts for a new parkrunner id, & dob
 *  2. Verifies runner matches row (with runner index)
 *  3. Removes the results sheet (by runner index)for the existing runner
 *  4. Removes runner's row from the 'Runners' sheet (impacting indices of others below)
 *  5. Renames the results sheets of runners with higher indices.
 */
function DeleteFamilyMember() {
  PromptForRunner(deleteCASE);    // default addCASE
  // callback to DoDeleteFamilyMember with [parkrunnerId,dob] from form
}

function DoDeleteFamilyMember(
  form = ['357161','1-Nov-87'])      // as a callback, values are passed as strings
{
  let [parkrunnerId,dob] = form;
  parkrunnerId = +parkrunnerId; // number expected for matching
  let runnersRange = gv.allRunnersSheet.getRange(
    gc.runnerNameCOLUMN+gc.runnersStartROW+":"+gc.parkrunnerIdCOLUMN
  );
  // Verify runner matches row (with runner index)
  let runnerIndex = FindRunnerIndex(runnersRange,parkrunnerId,dob);
  if (runnerIndex < 0)
    Logger.log('ERROR: Unable to find matching runner, '+parkrunnerId+
    '\n - with matching DoB, '+new Date(dob).toDateString());
  else {
    let forename = runnersRange.offset(runnerIndex,0).getValue();
    let runnerNameId = forename+'_'+runnerIndex;
    if (gc.debug)
      Logger.log('Deleting runner\'s result sheet, '+runnerNameId+
        ' (for parkrunner, '+parkrunnerId+')');
    let resultsSheet = gv.activeSpreadsheet.getSheetByName(runnerNameId);
    if (resultsSheet)
      gv.activeSpreadsheet.deleteSheet(resultsSheet);    // delete runner's results sheet
    gv.allRunnersSheet.deleteRow(gc.runnersStartROW+runnerIndex);   // 1 less row impacts others...
    let numRunnersImpacted = gv.allRunnersSheet.getLastRow()+1      // e.g. 3 = (60+1)...
      -(gc.runnersStartROW+runnerIndex);                           // ...-(3+55)
    Logger.log('Removed runner, '+parkrunnerId+
      ' entry with result sheet, '+runnerNameId+
      ' where '+numRunnersImpacted+' runner\'s Ids are impacted'); 
    let impactRange = runnersRange.offset(runnerIndex,0,numRunnersImpacted,1);
    for (let index=0; index<numRunnersImpacted; index++) {
      forename = impactRange.offset(index,0).getValue();
      let oldSheetName = forename+'_'+(runnerIndex+index+1);    // e.g. _56.._58
      let newSheetName = forename+'_'+(runnerIndex+index);      //  =>  _55.._57
      gv.activeSpreadsheet.getSheetByName(oldSheetName).setName(newSheetName);
      Logger.log('Runner\'s results sheet, '+oldSheetName+
        ' renamed as '+newSheetName);
    }
  }
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
  PromptForRunner();    // default addCASE
  // callback to DoAddFamilyMember with [parkrunnerId,dob,email] from form
}

async function DoAddFamilyMember(form) {
  let [parkrunnerId,dob,email] = form;
  parkrunnerId = +parkrunnerId;
  if (gc.debug) Logger.log('1. Prompt: '+form);
  return OpenChromeBrowser()
    .then(() => {   // allow time to open browser on server
      if (gc.debug) Logger.log('2. Open: '+parkrunnerId);
      return GetRunnerResultsPage(parkrunnerId)
        .then(resultsPage => {   // after load page in browser
          if (gc.debug) Logger.log('3a. Runner:'+parkrunnerId);
          let [runnerNames,gender] = GetRunnerDetails(resultsPage);
          if (gc.debug) Logger.log('3b. Details: '+runnerNames+' '+gender+' '+email+' '+dob+' '+' '+parkrunnerId);
          let runnerNameId = CreateRunnerResultsSheet(
            runnerNames,gender,  // to go into cols A & B, C
            email,dob,           // into cols.D & E (hidden for security, as also F..H)
            parkrunnerId);       // into col J (after derived age-category in col. I)
          if (gc.debug) Logger.log('4. Create sheet: '+runnerNameId);
          return ImportAllResultsForRunner(parkrunnerId,runnerNameId,resultsPage)
            .then(() => {
              let [runnerName,runnerIndex] = runnerNameId.split('_');
              Logger.log('Adding family member with their results: '+runnerName+'\t['+runnerIndex+']');
              LockCallerForwardsTo(threadBatchFN,'added',runnerNameId);
            });
        })
    })
    .catch(err => {
      Logger.log('ERROR: Add Family Member, '+parkrunnerId+'\n'+err);
      return CloseChromeBrowser();
    })
    .finally(() =>
      // return CloseChromeBrowser() // NOT yet until forked process is done!
      Logger.log('Ordinarily, preserve browser session until completed'));
}

/**
 * Adds the first Runner to the new family with basic details
 *  1. Creates new Results Sheet for first runner 
 *  2. Imports all the results for the new Runner (without Positions)
 *  3. Triggers the Batch process to append positions
 */
function AddFirstMember(
  parkrunnerId,runnerNames,resultsPage)
{   // assumes famSpreadsheetId now set and active Spreadsheet switched
  if (gc.debug) Logger.log('New spreadsheet Id: '+gv.activeSpreadsheetId);
  if (gc.debug) Logger.log('First parkrunner: '+parkrunnerId);
  const firstINDEX = 0;
  let runnerNameId = CreateRunnerResultsSheet(
    runnerNames,undefined,  // gender, with
    undefined,undefined,    // email & DoB already processed
    parkrunnerId,           // MUST not overwrite MAP formulae
    firstINDEX);    // update first row in new (active) Runners sheet
  return ImportAllResultsForRunner(parkrunnerId,runnerNameId,resultsPage)
    .then(() => {
      let [runnerName,runnerIndex] = runnerNameId.split('_');
      Logger.log('Forking new: '+runnerName+'\t['+runnerIndex+']');
      LockCallerForwardsTo(threadBatchFN,'forked',runnerNameId);
    })
    .catch(err => {
      Logger.log('ERROR: Add First Member, '+parkrunnerId+'\n'+err);
      return CloseChromeBrowser();
    })
    .finally(() =>
      // return CloseChromeBrowser() // NOT yet until forked process is done!
      Logger.log('Ordinarily, preserve browser session until completed'));
}

function CreateNewSpreadsheet(
  templateFolder = templateFOLDER,
  templateName = templateSPREADSHEET)
{
  let targetFolder = DriveApp.getFoldersByName(templateFolder).next();
  if (!targetFolder) {
    targetFolder = DriveApp.createFolder(templateFolder);
  }
  const files = DriveApp.getFilesByName(templateName);
  if (files.hasNext()) {
    return [targetFolder,files.next().getId()];
  } else {
    throw new Error('Template not found: '+templateName+' in '+templateFolder);
  }
}

async function InstantiateFamilySpreadSheet(
  clubType = clubTYPE,   // or Clubrunners
  runnerNames = ['Peter','WALLIS'],  // into cols A & B of 1st Runners row of new family Spreadsheet
  gender = 'Male',    // for col C
  email = '',         // for Col D
  dob = undefined,    // for Col E (hidden for security, as also F..H)
  parkrunnerId)       // for col J (after derived age-category in col. I)
{
  let [targetFolder,templateId] = CreateNewSpreadsheet(templateFOLDER,templateSPREADSHEET);
  if (gc.debug) Logger.log('Template Id: '+templateId);
  let templateFile = DriveApp.getFileById(templateId);
  if (gc.debug) Logger.log('Template File: '+templateFile);
  let familyName = runnerNames[1];
  let familySheetFile = familyName+' '+clubType;
  if (gc.debug) Logger.log('Family sheet: '+familySheetFile);
  let familySpreadsheet = templateFile.makeCopy(familySheetFile,targetFolder);  // temp
  familySpreadsheetId = familySpreadsheet.getId();
  if (gc.debug) Logger.log('Family spreadsheet Id: '+familySpreadsheetId);
  familySpreadsheet = SpreadsheetApp.openById(familySpreadsheetId); // real object
  // Now switch context completely hereafter...
  // ...with the exception, scripts remain from the originator 
  SpreadsheetApp.setActiveSpreadsheet(familySpreadsheet);   // flushes content?
  gv.activeSpreadsheet = familySpreadsheet;
  gv.activeSpreadsheetId = familySpreadsheetId;
  gv.allRunnersSheet = gv.activeSpreadsheet
    .getSheetByName(gc.runnersSheetNAME) // ...for 1st runner
  gv.allRunnersSheet.getRange(1,1)
    .setValue(familySheetFile);   // conveniently file name in A1
  gv.allRunnersSheet.getRange(gc.runnersStartROW,1,1,5)        // cols A..E ) 5 params
    .setValues([[...runnerNames,gender,email,dob]]);  //  MAP formulae undisturbed
  gv.allRunnersSheet.getRange(gc.runnersStartROW,gc.parkrunnerIdCOL)
    .setValue(parkrunnerId);
  return familySheetFile; // with GLOBAL gv.activeSpreadsheet & Id, and gv.allRunnersSheet
}

const spawnCASE = {
  title: 'Spawn New Family / Club',
  desc: 'This spawns a new family/club Spreadsheet that captures all results for that first '
        +'member. The surname of this parkrunner is taken as the name of the new Spreadsheet file '
        +'within which they will appear in the 1st row of the Runners sheet, with a separate sheet '
        +'(&lt;first name&gt;_0) containing their results, with detailed positions eventially '
        +'appearing (via a background batch process).',
  action: 'Spawn',
  handler: 'DoSpawnNewFamily'
};

/**
 * Spawns a new Family after prompting for details of the first runner
 *  1. Prompts for a new parkrunner id, etc.
 *  2. Opens the browser session
 *  3. Gets the Runner's Details on their Results Page
 *  4. Instantiates a new Family SpreadSheet
 *  4a.  Adds the first Row on empty Runners Sheet
 *  4b.  Establish the new spreadshhet as active henceforth
 *  5. Add First Member in new Spreadsheet (with results)
 */
function SpawnNewFamily() {
  PromptForRunner(spawnCASE); 
  // callback to DoSpawnNewFamily with [parkrunnerId,dob,email] from form
}

function DoSpawnNewFamily(
  form = ['21283','30-Jan-1969',undefined]
) {
  let [parkrunnerId,dob,email] = form;
  parkrunnerId = +parkrunnerId;
  if (gc.debug) Logger.log('1. Prompt: '+form);
  return OpenChromeBrowser()
    .then(() => {   // allow time to open browser on server
      if (gc.debug) Logger.log('2. Open: '+parkrunnerId);
      return GetRunnerResultsPage(parkrunnerId);
    })
    .then(resultsPage => {   // after load page in browser
      if (gc.debug) Logger.log('3a. Runner: '+parkrunnerId);
      let [runnerNames,gender] = GetRunnerDetails(resultsPage);
      if (gc.debug)
        Logger.log('3b. Details: '+runnerNames+' '+gender+' ('+email+' email) '+dob);
      return InstantiateFamilySpreadSheet(
        clubTYPE,
        runnerNames,gender, // into cols A & B, C of 1st row of new Runners instance
        email,dob,          // into cols.D & E (hidden for security, as also F..H)
        parkrunnerId)       // into col J (after derived age-category in col. I)
        .then(familySheetFile => [familySheetFile,runnerNames,resultsPage])
    })
    .then(([familySheetFile,runnerNames,resultsPage]) => {
      if (gc.debug) Logger.log('Family Spread sheet: '+familySheetFile);
      let [familyName,clubType] = familySheetFile.split(' ');
      if (gc.debug) {
        Logger.log('4. Instantiate 1st runner in: '+familyName+' ['+clubType+']');
      }
      AddFirstMember(parkrunnerId,runnerNames,resultsPage);
    })
    .catch(err => {
      Logger.log('ERROR: Spawn New Family for parkrunner, '+parkrunnerId+'\n'+err);
      return CloseChromeBrowser();
    })
    .finally(() =>
      // return CloseChromeBrowser() // NOT yet until forked process is done!
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
    .offset(gc.resultsStartROW-1,0,  // start row 3 = offset 2 (for title and header)
      resultsSheet.getLastRow()-(gc.resultsStartROW-1));   // e.g. 10 rows: 10-(3-1) = 8 results
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
  runnerNameId = 'Alan_2')    // default unit test (assumes precedent 
{
  const runnerIndex = parseInt(runnerNameId.split('_')[1]);
  const ageCatCELL = 'L3';  //runner may be older since 1st run but gender fixed
  let resultsSheet = gv.activeSpreadsheet.getSheetByName(runnerNameId);
  if (!resultsSheet) {
    Logger.log('ERROR: Unable to find results sheet, '+runnerNameId
      +' in current spreadsheet ('+gv.activeSpreadsheet+')');
    return false;
  }
  let runnerFullName = resultsSheet.getRange(gc.runnerNameCELL).getValue();
  var resultRange = FirstMatchRange(resultsSheet,gc.genderPosnCOL,"");  // 1st with unknown Gender posn
  if (resultRange) {   //skip this runner if all positions known
    // WARNING: Limit batch of results to catch up (recursively) because 6 mins max per App script!
    let lastResultRow = resultsSheet.getLastRow();
    let promises = [];
    let batchMore = gc.batchSizeMAX;
    while (batchMore) {
      let genderPositionKnown = resultRange.getCell(1,gc.genderPosnCOL).getValue();
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
    if (gc.debug) Logger.log('RunnerId: '+runnerNameId+'\t# of results is '+lastResultRow+' and reached result, '+resultRange.getLastRow());
    let remainToDo = lastResultRow-resultRange.getLastRow();
    var moreBatches = (remainToDo > 0);
    return Promise.all(promises).then(results => {
      let updatesApplied = results.filter(Boolean).length;
      let totalEvents = results.length;
      let morePositions = (moreBatches) ? ' with '+remainToDo+' more positions needed to catch up' : '';
      Logger.log('Runner: '+runnerFullName+'\tUpdates applied: '+updatesApplied
        +' (out of '+totalEvents+' in this batch)'+morePositions);
      if (updatesApplied === 0)
        throw new Error('ERROR: Batching aborted to avoid infinite looping because no updates');
      return moreBatches;
    });
  } else {
    if (gc.debug) Logger.log('Runner: '+runnerFullName+'\tNo (more) blank positions to catch up');
    return false;   // no more batches
  }
}

const threadBatchFN = recurseBatchFN ='BatchPositionsForRunner';  // threaded & potentially recursed
const lockINDEX = 'lockIndex';
const lockSPREADSHEET = 'lockSpreadsheet';     // in original or new family?

/**
 * Cleans up triggers and script properties after batching for each runner has completed
 *  @param {string} [thisScript=threadBatchFN] - script to clean up
 */
function CleanupBatch(
  thisScript = threadBatchFN) 
{
  var triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    Logger.log('Trigger handler cleaning: '+trigger.getHandlerFunction());
    if (trigger && trigger.getHandlerFunction() == thisScript)
      ScriptApp.deleteTrigger(trigger);
  });
  PropertiesService.getScriptProperties().deleteProperty(lockINDEX);
  PropertiesService.getScriptProperties().deleteProperty(lockSPREADSHEET);
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
  if (gc.debug) Logger.log('Finished: '+allDone);
  return allDone;
}

/**
 * Set this status of the runner in their results row
 *    @param {runnerIndex} indicates the runner to be marked as Positions all done
 *  @returns all of the runners Status is complete 
 */
function MarkRunnerPositionsDone(runnerIndex = 0) {
  var runnerStatusRange = gv.allRunnersSheet.getRange(
    gc.hasPosnsCOLUMN+gc.runnersStartROW+":"+gc.hasPosnsCOLUMN
  );
  runnerStatusRange.getCell(runnerIndex+1,1).setValue(true);
  let runnersStatus = runnerStatusRange.getValues().map(x => x[0]);
  if (gc.debug) {
    let changedStati = runnersStatus.map((status,index) =>
      ({index,status,changed: index === runnerIndex})
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
  let thisRunnerNameId = PropertiesService.getScriptProperties()
    .getProperty(lockINDEX);
  spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(lockSPREADSHEET);
  gv.activeSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  gv.activeSpreadsheetId = spreadsheetId;     // both Id & object set globally
  Logger.log('Batch processing of unique runner, '+thisRunnerNameId
    +' within spreadsheet, '+spreadsheetId);
  lock.releaseLock();   // unlock lock (unique for this node service)
  return thisRunnerNameId;
}

/**
 * Lock any further threaded/recursed for the uniques runner index until received
 */ 
function LockCallerForwardsTo(thisFunction,withReason,thisRunnerNameId) {
  var lock = LockService.getScriptLock();
  try {
    PropertiesService.getScriptProperties()
      .setProperty(lockINDEX,thisRunnerNameId);
    PropertiesService.getScriptProperties()
      .setProperty(lockSPREADSHEET,gv.activeSpreadsheetId);
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
function BatchPositionsForRunner() {
  const runnerNameId = UnlockCallerForwarded();
// function BatchPositionsForRunner(runnerNameId='Amy_20') {  // for gc.debug
  let [runnerName,runnerIndex] = runnerNameId.split('_');
// begin
  return OpenChromeBrowser()
  .then(() => { 
    runnerIndex = parseInt(runnerIndex);   // ensure a number 
    if (gc.debug)
      Logger.log('Runner: '+runnerName+' ['+runnerIndex+']\tSyncing positions...');
    return SyncPositionsPerRunner(runnerNameId);
  })
  .then((moreToDo) => { 
    if (moreToDo) {
      // Avoid limitation on triggers - clean as you go!
      ScriptApp.getProjectTriggers().forEach(trigger => {
        Logger.log('Trigger handler reduction: '+trigger.getHandlerFunction());
        if (trigger.getHandlerFunction() === 'BatchPositionsForRunner')
          ScriptApp.deleteTrigger(trigger);
      });
      Logger.log('Recursing runner: '+runnerName+'\t['+runnerIndex+']');
      LockCallerForwardsTo(recurseBatchFN,'recursed',runnerNameId);  
    } else {
      let runnersStatus = MarkRunnerPositionsDone(runnerIndex);
      if (AllPositionsDone(runnersStatus)) {
        CleanupBatch(threadBatchFN);
        return;
      } else  // loop back for next "unfinished" runner
        ScriptApp.newTrigger('CatchUpAllPositions')
          .timeBased()
          .after(1500)
          .create();
    }
  })
  .catch(err => {
    Logger.log('ERROR: Failed to trigger recursed batches: '+err);
  })
  .finally(() => {
    // Perhaps need to wait on closure before re-opening on next batch?
    return CloseChromeBrowser();   // open on each batch avoids unexpected session end
  });
}

/**
 * Triggers batch processing for runners needing position catch-up, always  in series
 */
function CatchUpAllPositions() {
  let runnersStatus = [];
  let runners = gv.allRunnersSheet
    .getRange(runnerNameCOLUMN+runnersStartROW+":"+runnerNameCOLUMN)
    .getValues().map(x => x[0]).filter(String);
  var runnersResults = gv.allRunnersSheet
    .getRange(gc.hasResultsCOLUMN+gc.runnersStartROW+":"+gc.hasResultsCOLUMN)
    .getValues().map(x => x[0]);
  // Ensure ALL threads use the same status so that closure is when done for ALL runners
  runnersStatus = gv.allRunnersSheet
    .getRange(gc.hasPosnsCOLUMN+gc.runnersStartROW+":"+gc.hasPosnsCOLUMN)
    .getValues().map(x => x[0]);
  // ONLY thread process for ONE valid runner initially, and let batching follow-on thereafter
  for (var [runnerIndex,runnerName] of runners.entries()) {
    if (runnersResults[runnerIndex]) {    // if runner has at least one result
      if (!runnersStatus[runnerIndex]) {  // ...and positions not already caught-up
        let runnerNameId = runnerName+'_'+runnerIndex;  // consistently unique for index and status
        Logger.log('Threading batch for a single runner: '+runnerName+' ['+runnerIndex+']');
        LockCallerForwardsTo(threadBatchFN,'threaded',runnerNameId);
        break; // batch the first incomplete runner only until done to avoid conflict
      } else if (gc.debug) 
        Logger.log('Positions up-to-date for runner: '+runnerName+' ['+runnerIndex+']');
    } else {
      if (!runnersStatus[runnerIndex]) {   // Since no results, then already caught-up
        runnersStatus = MarkRunnerPositionsDone(runnerIndex);   // avoids impact on not finishing
      }
    }
  }
  if (AllPositionsDone(runnersStatus)) {
    CleanupBatch(threadBatchFN);
    return;
  }
}

/**
 * Accept cookies on all country Domains
 */
async function AcceptCookies() {
  const acceptCookiesViaBrowserURL = gc.browserURL+'/acceptCookies';
  try {
    var response = await UrlFetchApp.fetch(acceptCookiesViaBrowserURL,
      {muteHttpExceptions:true,timeout:30000});
    return response.getContentText();
  } catch (err) {
    Logger.log(err);
    return err;
  }
}
