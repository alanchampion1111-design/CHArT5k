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
/   1.  ImportResultForEachRunner (optional event date of Parkrun)
/         OpenChromeBrowser ->
//        Loop for each member runner...
/           >CopyResultForRunner (latest or dated?)->
/             >GetRunnerResultsPage
/               >AccessPage
/             GetResultRow
//          When latest is a new result...
/             >PasteResultForRunner ->
//              Loop for each cell...
/                 CleanValue (apply links as hyperlinks)
/               FormatDate (may be redundant)
/               AppendResultRow
/             >AppendPositionsForResult ->
/               GetResultUrl (assume always parkrun)
/               GetDomainGender  (for language-based filter)
/               ExtendRange
/               >AssessPositions (for one runner)
/               IncludePositions
/         >CloseChromeBrowser
/
/   2.  CatchUpAllPositions
/         OpenChromeBrowser ->
//        Loop for each member runner...
/           LockCallerForwardsTo ->
/
/   2a. BatchPositionsForRunner (threaded in parallel?)
/         UnlockCallerForwarded
/         OpenChromeBrowser ->
/           SyncPositionsPerRunner -> (in batches of 10 results)
/             FirstMatchRange (based on unknown Gender position)
//            When positions not previously updated...
//              Loop for each runner's result (per batch)...
/                 AppendPositionsForResult ->
/                   GetResultUrl
//                  When event is a Parkrun...
/                     GetDomainGender (for language-based filter)
/                     >AssessPositions (for one runner)
/                     IncludePositions
/                       Trigger forked task, BatchPositionsForRunner
/
/   2b. BatchPositionsForRunner (recurse if more to sync):
/         UnlockCallerForwarded
/         When all positions applied...
/           >CloseChromeBrowser
/  -------------------------------------------------------------------------
*/

/**
 * For the purpose of this trial, the following session & connection functions rely on
 * The asynchronous functions include:
 *    OpenChromeBrowser
 *    AccessPage
 *    CloseChromeBrowser
 */

async function OpenChromeBrowser() {
  const initBrowserURL = browserURL+'/initBrowser';
  try {
    var response = await UrlFetchApp.fetch(initBrowserURL);
    browserSession = response.getContentText();   // WS Endpoint lingers for upto 30 minutes
    Utilities.sleep(1000); // in case image not cached, avoids Error: Requesting main frame too early!
    Logger.log('Session: '+browserSession);
    return true;
  } catch (err) {
    Logger.log(err);
    return false;
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
  var resultsSheet = activeSpreadsheet.getSheetByName(runnerNameId);
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
    // TODO redundant since opened on each batch, with managable batch size (<20)
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
    // Pending outcome of proposal to Parkrun site: "Replace the Run # with Gender Position?
  // Alternatively, the number of runners in the event may replace the redundant run number
  resultRange.offset(0,GenderPosnCOL-1,1,3).setValues([[gcPosn, acPosn, agPosn]]);
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
    'en': { M: 'Male', W: 'Female' },
    'nl': { M: 'Man', W: 'Vrouw' },
    'de': { M: 'Männlich', W: 'Weiblich' },
    'it': { M: 'Uomo', W: 'Donna' },
    'fr': { M: 'Homme', W: 'Femme' },
    'da': { M: 'Mand', W: 'Kvinde' },
    'fi': { M: 'Mies', W: 'Nainen' },
    'no': { M: 'Mann', W: 'Kvinne' },
    'pl': { M: 'Mężczyzna', W: 'Kobieta' },
    'sv': { M: 'Man', W: 'Kvinna' },
    'jp': { M: 'Otoko', W: 'Onna' }
  };
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
    'no': 'no',
    'pl': 'pl',
    'se': 'sv',
    'jp': 'jp',
    'ca': 'en', // or 'fr' for French-speaking Canada
    'sg': 'en',
    'nz': 'en',
    'my': 'en', // or 'ms' for Malay
    'na': 'en',
    'lt': 'lt', // add Lithuanian translations to languageMAP
    'es': 'en' // add Eswatini, assuming English
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
    var runners = allRunnersSheet.getRange(
      runnerNameCOLUMN+runnersStartROW+":"+runnerSurnameCOLUMN
    ).getValues().filter(String);
    // Process each runner in parallel BEFORE closing after ALL runners done!
    return Promise.all(runners.map(function(runner,index) {
      let runnerName = runner[0];     // col A for first name
      let runnerNameId = runnerName+'_'+index;
      let parkrunnerId = allRunnersSheet.getRange(
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
            let runnerFullName = runner.join(' ');  // from col A & B
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

function AppendAllResults(
  runnerNameId,
  linksCount,allFormulae,
  valuesCount,allValues)
{
  var resultsSheet = activeSpreadsheet.getSheetByName(runnerNameId);
  let numResults = allFormulae.length;    // = allValues.length
  resultsSheet.getRange(resultsStartROW,1,numResults,linksCount)
    .setFormulas(allFormulae);
  resultsSheet.getRange(resultsStartROW,linksCount+1,numResults,valuesCount)
    .setValues(allValues);
  resultsSheet.getRange(resultsStartROW+1,PBtickBoxCOL,numResults-1,1)   // column beyond paste
    .clear({contentsOnly: true})  // complement to MAP prevents FALSE 
    .insertCheckboxes();          // c/fwd tickbox restored
}

function PasteAllResultsForRunner(
  runnerNameId = 'Sarah_17',
  allResults)
{
  const dateINDEX = 1; // column B
  let linksCount = 3;
  let valuesCount = 4;
  let allFormulae = [];
  let allValues = [];
  // Paste in reverse order from parkrun site => earliest result first
  allResults.reverse().forEach(thisResult => {
    thisResult = thisResult.map(CleanValue);        // extracts hyperLinks...
    thisResult[dateINDEX] = FormatDate(thisResult[dateINDEX],dateFORMAT);
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
  if (debug) Logger.log('Import All from parkrunner, '+parkrunnerId+' to results sheet, '+runnerNameId);
  return CopyResultForRunner(parkrunnerId,'ALL',thisPage)
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
    rangeRow = allRunnersSheet.appendRow(
      [...runnerNames,        // A..B
      gender,email,dob,       // C..E
      undefined,undefined,undefined,undefined,  // F..I since derived categories
      parkrunnerId,null,null] // J..L parkrun + derived tick boxes: has results. has positions
    );
    runnerIndex = allRunnersSheet.getLastRow()-runnersStartROW;
  } else {
    // Assume details already on Runners Sheet for first runner
  }
  // Create runner's results sheet if it doesn't exist
  let [runnerName,runnerSurname] = runnerNames;
  let runnerNameId = runnerName+'_'+runnerIndex;
  let templateResults = activeSpreadsheet.getSheetByName(templateNAME);
  let newResultsSheet = templateResults.copyTo(activeSpreadsheet).setName(runnerNameId);
  // ensure the content of the new sheet is unique
    let runnerFullName = runnerNames.join(" ");
    const titleNameCELL = "A1";
    const allRunnersGID = allRunnersSheet.getSheetId();
    newResultsSheet.getRange(titleNameCELL)
      .setFormula('=HYPERLINK("#gid='+allRunnersGID+'","'+runnerFullName+'")');
  // return to Runners sheet and add the link back and the fast results link
    let newResultsGid = newResultsSheet.getSheetId();
    let parkrunnerResultsUrl = parkrunnerURL+parkrunnerId+allForONE;
    rangeRow = allRunnersSheet.getLastRow();
    allRunnersSheet.getRange(rangeRow,1)     // Col. A
      .setFormula('=HYPERLINK("#gid='+newResultsGid+'","'
        +runnerName+'")');
    allRunnersSheet.getRange(rangeRow,parkrunnerIdCOL)     // Col. J    
      .setFormula('=HYPERLINK("'+parkrunnerResultsUrl+'","'
        +parkrunnerId+'")');
    if (runnerIndex != 0) {
      let numCols = allRunnersSheet.getLastColumn();
      allRunnersSheet.getRange(rangeRow-1,1,1,numCols)
        .copyFormatToRange(allRunnersSheet,1,numCols,rangeRow,rangeRow);
    }
  return runnerNameId;    // name of the newly created sheet
}

/*
/  Hierarchy for two use-cases of a new runner:
/
/   1.  AddFamilyMember (as a member of your family/club)
/         PromptNewRunner(addCASE)-> to get parkrun id, Dob,..
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
/         PromptNewRunner(spawnCASE)-> to get parkrun id, Dob,..
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
/             FormatDate (TODO: potentially redundant)
/             PrepareResultRow
/             AppendAllResults
/         LockCallerForwardsTo-> (triggers)
//          Trigger task, BatchPositionsForRunner...
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
  if (debug) Logger.log ('3b. Name: '+runnerFullName);
  let runnerNames = runnerFullName.split(' ');
  if (runnerNames.length > 2) {   // Keep any middle name with surname 
    runnerNames = [runnerNames[0],runnerNames.slice(1).join(' ')];
  }
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
  return [runnerNames,gender]; 
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
  return OpenChromeBrowser()
    .then(() => {   // allow time to open browser on server
      if (debug) Logger.log('2. Open: '+parkrunnerId);
      return GetRunnerResultsPage(parkrunnerId);
    })
    .then(resultsPage => {   // after load page in browser
      if (debug) Logger.log('3a. Runner:'+parkrunnerId);
      let [runnerNames,gender] = GetRunnerDetails(resultsPage);
      if (debug) Logger.log('3b. Details: '+runnerNames+' '+gender+' '+email+' '+dob+' '+' '+parkrunnerId);
      let runnerNameId = CreateRunnerResultsSheet(
        runnerNames,gender,  // to go into cols A & B, C
        email,dob,           // into cols.D & E (hidden for security, as also F..H)
        parkrunnerId);       // into col J (after derived age-category in col. I)
      if (debug) Logger.log('4. Create sheet: '+runnerNameId);
      return ImportAllResultsForRunner(parkrunnerId,runnerNameId,resultsPage)
      	.then(() => {
          let [runnerName,runnerIndex] = runnerNameId.split('_');
          Logger.log('Adding family member with their results: '+runnerName+'\t['+runnerIndex+']');
          LockCallerForwardsTo(threadBatchFN,'added',runnerNameId);
        });
    })
    .catch(err => {
      Logger.log('ERROR: Add Family Member, '+parkrunnerId+'\n'+err);
      CloseChromeBrowser();
    })
    .finally(() =>
      // CloseChromeBrowser() // NOT yet until forked process is done!
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
  if (debug) Logger.log('New spreadsheet Id: '+activeSpreadsheetId);
  if (debug) Logger.log('First parkrunner: '+parkrunnerId);
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
      CloseChromeBrowser();
    })
    .finally(() =>
      // CloseChromeBrowser() // NOT yet until forked process is done!
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
  if (debug) Logger.log('Template Id: '+templateId);
  let templateFile = DriveApp.getFileById(templateId);
  if (debug) Logger.log('Template File: '+templateFile);
  let familyName = runnerNames[1];
  let familySheetFile = familyName+' '+clubType;
  if (debug) Logger.log('Family sheet: '+familySheetFile);
  let familySpreadsheet = templateFile.makeCopy(familySheetFile,targetFolder);  // temp
  familySpreadsheetId = familySpreadsheet.getId();
  if (debug) Logger.log('Family spreadsheet Id: '+familySpreadsheetId);
  familySpreadsheet = SpreadsheetApp.openById(familySpreadsheetId); // real object
  // Now switch context completely hereafter...
  // ...with the exception, scripts remain from the originator 
  SpreadsheetApp.setActiveSpreadsheet(familySpreadsheet);   // flushes content?
  activeSpreadsheet = familySpreadsheet;
  activeSpreadsheetId = familySpreadsheetId;
  allRunnersSheet = activeSpreadsheet
    .getSheetByName(runnersSheetNAME) // ...for 1st runner
  allRunnersSheet.getRange(1,1)
    .setValue(familySheetFile);   // conveniently file name in A1
  allRunnersSheet.getRange(runnersStartROW,1,1,5)        // cols A..E ) 5 params
    .setValues([[...runnerNames,gender,email,dob]]);  //  MAP formulae undisturbed
  allRunnersSheet.getRange(runnersStartROW,parkrunnerIdCOL)
    .setValue(parkrunnerId);
  return familySheetFile; // with GLOBAL activeSpreadsheet & Id, and allRunnersSheet
}

const spawnCASE = {
  title: 'Spawn New Family / Club',
  desc: 'This spawns a new family/club Spreadsheet that captures all results for that first '
        +'member. The surname of this parkrunner is taken as the name of the new Spreadsheet file'
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
  PromptNewRunner(spawnCASE); 
  // callback to DoSpawnNewFamily with [parkrunnerId,dob,email] from form
}

function DoSpawnNewFamily(
  form = ['21283','30-Jan-69',undefined]
) {
  let [parkrunnerId,dob,email] = form;
  if (debug) Logger.log('1. Prompt: '+form);
  const dobRegex = /^\d{2}-[A-Za-z]{3}-\d{2}$/;
  if (!dobRegex.test(dob)) {
    throw new Error('ERROR: Invalid date format - use dd-Mmm-YY');
    // otherwise coerce to this format
  }
  return OpenChromeBrowser()
    .then(() => {   // allow time to open browser on server
      if (debug) Logger.log('2. Open: '+parkrunnerId);
      return GetRunnerResultsPage(parkrunnerId);
    })
    .then(resultsPage => {   // after load page in browser
      if (debug) Logger.log('3a. Runner: '+parkrunnerId);
      let [runnerNames,gender] = GetRunnerDetails(resultsPage);
      if (debug)
        Logger.log('3b. Details: '+runnerNames+' '+gender+' ('+email+' email) '+dob);
      return InstantiateFamilySpreadSheet(
        clubTYPE,
        runnerNames,gender, // into cols A & B, C of 1st row of new Runners instance
        email,dob,          // into cols.D & E (hidden for security, as also F..H)
        parkrunnerId)       // into col J (after derived age-category in col. I)
        .then(familySheetFile => [familySheetFile,runnerNames,resultsPage])
    })
    .then(([familySheetFile,runnerNames,resultsPage]) => {
      if (debug) Logger.log('Family Spread sheet: '+familySheetFile);
      let [familyName,clubType] = familySheetFile.split(' ');
      if (debug) {
        Logger.log('4. Instantiate 1st runner in: '+familyName+' ['+clubType+']');
      }
      AddFirstMember(parkrunnerId,runnerNames,resultsPage);
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
  runnerNameId = 'Alan_2')    // default unit test (assumes precedent 
{
  const runnerNameCELL = 'A1';
  const ageCatCELL = 'L3';  //runner may be older since 1st run but gender fixed
  const dateINDEX = 1;      // index for column B
  const genderPosnCOL = 9;  // column I is the Gender position (if present already done)
  let resultsSheet = activeSpreadsheet.getSheetByName(runnerNameId);
  if (!resultsSheet) {
    Logger.log('ERROR: Unable to find results sheet, '+runnerNameId
      +' in current spreadsheet ('+activeSpreadsheet+')');
    return false;
  }
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
const lockSPREADSHEET = 'lockSpreadsheet';     // in original or new family?
const runnerNameCOLUMN = "A";     // Runners name in Column A 
const runnerSurnameCOLUMN = "B";  // Runners surname in Column B 
const hasResultsCOLUMN = "K";     // ...results exist (D3:D), with Parkrunner Id in col J
const hasPosnsCOLUMN = "L";       // ...has Positions up-to-date (I3:I) based on GenderPosn

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
  PropertiesService.getScriptProperties().deleteProperty(lockSPREADSHEET);
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
  var runnerStatusRANGE = allRunnersSheet.getRange(hasPosnsCOLUMN+runnersStartROW+":"+hasPosnsCOLUMN);
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
  let thisRunnerNameId = PropertiesService.getScriptProperties()
    .getProperty(lockINDEX);
  spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(lockSPREADSHEET);
  activeSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
  activeSpreadsheetId = spreadsheetId;     // both Id & object set globally
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
    lock.waitLock(12000);  // max time before parallel threads may continue
    PropertiesService.getScriptProperties()
      .setProperty(lockINDEX,thisRunnerNameId);
    PropertiesService.getScriptProperties()
      .setProperty(lockSPREADSHEET,activeSpreadsheetId);
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
  let [runnerName,runnerIndex] = runnerNameId.split('_');
// begin
  return OpenChromeBrowser()
  .then(() => { 
    runnerIndex = parseInt(runnerIndex);   // ensure a number 
    if (debug)
      Logger.log('Runner: '+runnerName+' ['+runnerIndex+']\tSyncing positions...');
    return SyncPositionsPerRunner(runnerNameId);
  })
  .then((moreToDo) => { 
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
  .finally(() => {
    // CloseChromeBrowser() // NOT yet until status for every
    Logger.log('Ordinarily, preserve browser session until completed');
    // Avoid limitation on triggers - clean as you go!
    let triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => {
      if (trigger.getTriggerSource() === ScriptApp.TriggerSource.CLICK
          && !trigger.isEnabled()) {
        ScriptApp.deleteTrigger(trigger);
      }
    });
  })
}

/**
 * Retrospectively update Positions for each runner via separately triggered processes
 * in parallel, where batching also may be necessary to avoid exceeding 6 seconds max per set.
 */
function CatchUpAllPositions() {
  let runnersStatus = [];
  return OpenChromeBrowser().then(() => {   // Browser always launched beforehand...
    let runners = allRunnersSheet.getRange(runnerNameCOLUMN+runnersStartROW+":"+runnerNameCOLUMN)
      .getValues().map(x => x[0]).filter(String);
    var runnersResults = allRunnersSheet.getRange(hasResultsCOLUMN+runnersStartROW+":"+hasResultsCOLUMN)
      .getValues().map(x => x[0]);
    // Ensure ALL threads use the same status so that closure is when done for ALL runners
    runnersStatus = allRunnersSheet.getRange(hasPosnsCOLUMN+runnersStartROW+":"+hasPosnsCOLUMN)
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
