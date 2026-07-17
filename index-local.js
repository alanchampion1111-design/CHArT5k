// Multiple steps to do:
//    0. Verify base image is in Artifact Registry (done)
//    1. Verify triggers gets latest sources from GitHub including this index.js file (done)
//    2. Verify build image uses Docker to install Chrome (done)
//    3. Verify Chrome browser works directly on server after build (done)
//    4. Verify server side, thisBrowser activated by client from Google Spreadsheet app (done)
//    5. Verify sample page retrieved ok and that thisBrowser and thisPage persists (done)
//    6. Upload profile/certificates for access to www.parkrun.owaitfoultsrrg.uk (done)
//    7. Verify allowed to load content for www.parkrun.org.uk (done)
//    8. Verify stealth access to individual parkrunner results table (tbd - although disallowed)

// const functions = require('@google-cloud/functions-framework');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());
// const userAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/150.0.7871.125 Safari/537.36';
const url = require('url');
let thisBrowserWSEp;  // re-use same browser session on server   
let thisPageId;       // re-use same page per session
const http = require('http');
const HOST = '127.0.0.1';    // equivalent to localhost
const PORT = 36007;
const browserURL = 'http://'+HOST+':'+PORT;
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');
const parkrunURL = 'https://www.parkrun.org.uk';    // TODO: unltimately depends on owner's native site
const parkrunnerURL = parkrunURL+'/parkrunner/';
const chromePATH = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
const myUserDataDir = 'C:\\CHArT5k-Puppet\\userDataDir';
// Assumes each Parkrun domain certificates have been exported (from normal use) and held in GitHub directly
// TODO: better to place in a certs subfolder (and limit access?)
const allParkrunCERTS =
  './www.parkrun.org.uk.pem,'+
  './www.parkrun.com.pem,'+
  './www.parkrun.co.nl.pem,'+
  './www.parkrun.com.de.pem,'+
  './www.parkrun.com.au.pem,'+
  './www.parkrun.ca.pem,'+
  './www.parkrun.jp.pem';

let prevPage;         // retain thisPage on weekly import (minimal cache => single page)
let prevFilterUrl;    // recall previous event URL determines re-use of retained prevPage
let browserTimeout;   // for browser session
let browserTimer;
let cachedPages = {};    // stores separate open URL pages when caching
const sessionMINS = 60;  // browser session timeout about one hour
const launchSECS = 50;  // initial first page load timeout
const pageSECS = 3;     // Assume 10 seconds BETWEEN page accesses on parkrun site relies on stealth mode?
const minRunnerTableCOUNT = 3;  // 3 for a 5k runner
const loadSECS = 30;            // max time to load each runner's results page
const minClubTableCOUNT = 1;    // 1..10+ tables for event locations for a family/club? 
const loadDetailSECS = 40;      // max time to load any event results page or global site consolidated club results page
let isLaunching = false; // PREVENT concurrent re-launch spikes inside loadUrl

/**
 *  Launches a headless Chrome browser with specified session limit.
 *    @param {number} [sessionLimit default 5] - Session limit in minutes
 *  @returns {Promise<void>}
 *  @sideeffect leaves the browser connected and returns a presistent WS endpoint for re-use
 */
let cloudBrowser = async (
  sessionMins = sessionMINS) =>
{
  browserTimeout = sessionMins*60*1000;
  var thisBrowser = await puppeteer.launch({  // variable delay if image not cached?
    headless: true,
    executablePath: chromePATH,
    userDataDir: myUserDataDir, // recall cookies and certs
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--cert='+allParkrunCERTS,
      '--verbose',
      '--lang=en-GB'    //  ensures the date formats appear as dd/mm/yyyy
    ],
    timeout: launchSECS*1000,       // max launch time
    // detached: true,         // ensure session with puppeteer persists after initial launch
    // ignoreHTTPSErrors: true
  });
  // Set a timer to close the browser by default after the timeout
  browserTimer = setTimeout(async () => {
    try {
      console.warn('WARNING: Terminating browser due to timeout:',browserTimeout);
      await thisBrowser.close();
    } catch (err) {
      console.error('ERROR: Terminating browser on timeout:',err);
    } finally {
      thisBrowserWSEp = null;
      thisPageId = null;
      prevPage = undefined;
      prevFilterUrl = undefined;
      clearTimeout(browserTimer);
    }
  }, browserTimeout);
  thisBrowserWSEp = thisBrowser ? thisBrowser.wsEndpoint() : null;  // return to client (although also global)
  console.log('Retained browser WS Endpoint:',thisBrowserWSEp);
  cachedPages = {};  // applicable to filterUrl only for shared events on same date
  try {
    var thisPage = await thisBrowser.newPage();
    if (thisPage) {
      thisPageId = await thisPage.target()._targetId;
      console.log('Retained page ID,',thisPageId);
      thisPage.setDefaultTimeout(loadSECS*1000);  // Set the timeout for loading the page
      await thisPage.goto('about:blank',{waitUntil: 'domcontentloaded'});    // Verify the browser is ready
      var content = await thisPage.content();    // always ensure page is fully loaded
      console.log('Blank page fully loaded');
    } else {
      console.warn('WARNING: Potentially failed to retain page ID,',thisPageId);
    }
  } catch (err) {
    console.error('ERROR: Getting page ID:', err);
  }
}

function forceKillBrowser() {
  try {
    const cmd = `wmic process where "commandline like '%C:\\\\CHArT5k-Puppet\\\\userDataDir%'" get ProcessId`;
    const output = execSync(cmd).toString();
    const pids = output.split('\n').slice(1).filter(line => line.trim() !== '');
    pids.forEach(pid => {
      console.log(`Killing process tree for PID: ${pid}`);
      exec(`taskkill /F /T /PID ${pid.trim()}`);    // /F = Force, /T = Tree (kills children), /PID = Target
    });
    let lockFile = path.join(myUserDataDir,'SingletonLock');
    if (fs.existsSync(lockFile)) {
      fs.unlinkSync(lockFile);
      console.log('Successfully removed orphaned SingletonLock.');
    }
  } catch (e) {
    console.log("Cleanup complete or no processes found.");
  }
}

/**
 * Kills the current browser instance and resets globals in case still active.
 *  @returns {Promise<boolean>} true if browser terminated, otherwise already terminated (perhaps error logged)
 */
let killBrowser = async () => {
  try {
    if (thisBrowserWSEp) {
      var thisBrowser = await puppeteer
        .connect({ browserWSEndpoint: thisBrowserWSEp });
      if (thisBrowser && thisPageId) {
        var thisPage = (await thisBrowser.pages())
          .find(page => page.target()._targetId === thisPageId);
        if (thisPage) {
          await thisPage.close();
          console.log('Page closed successfully - Page Id:',thisPageId);
        } else {
          console.warn('WARNING: Page previously closed or timed out - Page Id: ',thisPageId);
        }
      }
      let connectedBrowser = false;
      try {
        connectedBrowser = typeof thisBrowser.isConnected === 'function' 
          ? thisBrowser.isConnected() 
          : false;
        if (connectedBrowser) {
          await thisBrowser.close();
          console.log('Browser terminated successfully - WS endpoint:',thisBrowserWSEp);
          return true;
        } else {    // force kill any process to unlock
          console.warn('WARNING: Browser unreachable, attempting force-cleanup...');
          forceKillBrowser();
          return true; // Return success since forced the environment to clear
        }
      } catch (err) {
        console.warn('WARNING: Browser connection lost, already closed');
        return true;  // as if closed
      }  
    } else {
      console.warn('WARNING: Browser previously terminated - WS endpoint:',thisBrowserWSEp);
      return true;  // as if closed
    }
  } catch (err) {
    console.error('FATAL: Failed to close page and/or terminate browser:',err);
    return false;  // unsafe to continue if unknown reason
  } finally {  // executed in all cases, even before the returns
    thisBrowserWSEp = null;
    thisPageId = null;
    prevPage = undefined;
    prevFilterUrl = undefined;
    console.log('INFO: Expected browser page also closed :',thisPageId);
    clearTimeout(browserTimer);
  }
}

/**
 * Entry point: initialises browser with a new WebSocket endpoint (clearing existing if such exists)
 *    @param {Object} _ - void request
 *    @param {Object} res - HTTP fetch response to send
 *  @sideeffect - preserve the browser WS endpoint and reusable page ID for re-use
 */
exports.initBrowser = async (_,res) => {
  let isEnvironmentReady = await killBrowser();  // fresh start whenever re-open...
  if (!isEnvironmentReady) {
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('FATAL: Browser environment blocked. Manual intervention required.');
    return; 
  }
  try {
    await cloudBrowser(60);  // Launched ok, with browser active in background
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end(thisBrowserWSEp);
  } catch (err) {
    console.error('ERROR: Failed to initialise browser:',err);
    // consider a relaunch with args, --pull from Docker if image is not properly cached!
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('ERROR: Failed to initialise browser, ' + err);
  } finally {
    // NEVER disconnect because this loses the puppeteer Stealth (plugin) setting!
    // await thisBrowser.disconnect();
    console.log('Completed (attempt at) launching browser');
  }
};

/**
 * Entry point: stops the browser. Returns 200 if terminated, 204 if already terminated.
 *    @param {Object} _ - void request
 *    @param {Object} res - HTTP fetch response to send 200 if terminated, otherwise 204 if not so or not needed.
 */
exports.stopBrowser = async (_, res) => {
  let terminatedOk = await killBrowser();
  if (terminatedOk) {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end('INFO: Browser terminated successfully');
    // res.status(200).send('Browser terminated successfully');
  } else {
    res.writeHead(204, { 'Content-Type': 'text/plain' });
    res.end('WARNING: Browser previously terminated or timed out');
    // res.status(204).send('WARNING: Browser previously terminated or timed out');
  }
};

/** ______________________________________________________________________________________
/
  Functions heirarchy follows for this block used maintain 

    exports.getrUrl
      -> loadUrl

    exports.filterUrl
      -> loadUrl
      sortAgeGrade
        waitForResults
        sortPositions
        getRunnerNames
        getMatchName
        sortPositions (to unsort)
      filterCategory* (for agegroup & gender)
        waitForResults
        filterPositions
        getRunnerNames
        getMatchName
        removeCategory
/ ________________________________________________________________________
*/

/**
 * Loads URL in Puppeteer and waits for page load
 *   @param {string} thisUrl - URL to load
 *   @param {number} tableCount [default 3] -  (-1 meansd returns thisPage)
 *   @param {number} timeSecs - max timeout (in secs) to load (with table?)
 *   @param {boolean} caching - if true, page remains open and is to be (or may have been) cached (when tableNum is 0)
 *  @returns the HTML content when table # is loaded, or thisPage object for alternative detailed searching
 */
async function loadUrl(thisUrl,
  tableCount = 3,  // default for getUrl for a runner's 5k results, -1 for filterUrl on an event results
  timeSecs = loadSECS,
  caching = false)
{
  // console.log('Reconnecting to browser WS Endpoint:',thisBrowserWSEp,'with same page ID,',thisPageId);
  try {
    var timeMax = timeSecs*1000;
    const selectResultsTABLE = 'table[id*="results"]';  // 3 tables with id="results"
    if (!thisBrowserWSEp && !isLaunching) {  // Auto-heal if endpoint was wiped by a timeout or container spin-up
      console.warn('WARNING: Persistent browser NOT found. Re-launching internally...');
      isLaunching = true;
      await cloudBrowser(60);
      isLaunching = false;
    } else if (!thisBrowserWSEp) {
      throw new Error('Persistent browser is currently launching elsewhere!');
    }
    console.log('Persistent browser found,',thisBrowserWSEp,'with ongoing timeout:',browserTimeout);
    var thisBrowser;
    try {
      thisBrowser = await puppeteer
        .connect({browserWSEndpoint: thisBrowserWSEp, timeout: 10000});     // actually reconnects 
    } catch (connectErr) {
      // Catch hidden socket timeouts or closed browser instances cleanly
      console.error('ERROR: Connection to WS Endpoint failed. Re-building browser session:', connectErr);
      clearTimeout(browserTimer);
      isLaunching = true;
      await cloudBrowser(60);
      isLaunching = false;
      thisBrowser = await puppeteer
        .connect({browserWSEndpoint: thisBrowserWSEp, timeout: timeMax});
    }
    var thisPage;
    if (caching) {  // CAUTION: Caching common event is slower with new page overheads (rely on parkrun caching instead?)
      thisPage = await thisBrowser.newPage();
      await thisPage.setDefaultTimeout(timeMax);    // for individual queries?
    } else {  // re-use
      thisPage = (await thisBrowser.pages())
        .find(page => page.target()._targetId === thisPageId);
      if (!thisPage) {
        console.warn('WARNING: Persistent page missing from warm browser. Spawning fresh target context...');
        thisPage = await thisBrowser.newPage();
        thisPageId = await thisPage.target()._targetId;
        // await thisPage.setUserAgent(userAgent);
      }
    }
    await thisPage.setDefaultTimeout(timeMax); // Ensure active page timeout matches current transaction scope
    await thisPage.goto(thisUrl,{waitUntil: 'domcontentloaded',timeout: timeMax});
    if (tableCount < 0) {
      console.log('Event results page loaded for detailed sorting/filtering of URL,\n',thisUrl);
      return thisPage;
    } else {  // 5k page assumed with full history + PB column (Gender position missing!)
      await thisPage.waitForFunction((tableCount,selectResultsTABLE) =>
        document.querySelectorAll(selectResultsTABLE).length >= tableCount,
        {timeout: timeMax},tableCount,selectResultsTABLE);
      console.log('Runner results with ['+tableCount+'] tables loaded for URL,\n',thisUrl);
      return await thisPage.content();   // when page content is fully loaded
    }
  } catch (err) {
    isLaunching = false;
    console.error('ERROR: Failed to retrieve results page:',err);
    throw err;
  }
}

/**
 *  Gets the entire HTML content for a URL, typically all of results for a specific parkrunner 
 *    @param {string} url - e.g. https://www.parkrun.com/results/consolidatedclub/?clubNum=2548&eventdate=2026-07-03
 *  Note that the URL passes through any &eventdate parameter (if specified for consolidated results)
 *  @returns {string} as HTML content
 */
exports.getUrl = async (req,res) => {
  // Default for testing sample parkrunner 5k results OR sample consolidated results 
  // let thisUrl = req.query?.url || 'https://www.parkrun.org.uk/parkrunner/777764/5k/';
  var thisUrl = req.query?.url || 'https://www.parkrun.com/results/consolidatedclub/?clubNum=2548';
  let eventDate = req.query?.eventdate || null;    // '&eventdate=2026-07-03';
  if (eventDate)
    thisUrl = thisUrl+'&eventdate='+eventDate;
  let clubResults = thisUrl.includes('consolidatedclub');
  let [minTableCount,timeoutSecs] = (clubResults)
    ? [minClubTableCOUNT,loadDetailSECS] : [minRunnerTableCOUNT,loadSECS]; 
  try {
    var content = await loadUrl(thisUrl,minTableCount,timeoutSecs);
    res.writeHead(200, { 'Content-Type': 'application/json' });
    res.end(content);
    // res.status(200).send(content);
  } catch (err) {
    console.error(err);
    res.writeHead(500, { 'Content-Type': 'application/json' });
    res.end('ERROR: Failed to load URL, '+thisUrl,' - ',err);
    // res.status(500).send('ERROR: Failed to load URL, '+thisUrl);
  } finally {
    // AVOID disconnect because this loses the puppeteer Stealth (plugin) setting!
    // await thisBrowser.disconnect(); 
  }
}

/**
 *  Extracts runner names from results table on page
 *    @param {Page} thisPage - Puppeteer page object containing a table of results
 *  @returns {array} of runner names from the results table
 */
/*
  <table class="ResultsTable js-ResultsTable Results-table--compact">
    <tbody class="js-ResultsBody">
      <tr class="Results-table-row" data-name="James HARWOOD" data-agegroup="VM35-39" ... data-gender="Male" data-position="1" ...>
*/
async function getRunnerNames(thisPage) {
  const resultsTABLE = 'tr.Results-table-row';
  // await thisPage.waitForSelector(resultsTABLE);  // already follows waaitForResults that does this
  return await thisPage.$$eval(resultsTABLE,
    rows => rows.map(row => row.getAttribute('data-name'))
  );
}

/**
 *  Gets position of match name from page
 *    @param {array} names - list of runners by name
 *    @param {string} name - exists within array
 *  @returns {number} position - matching name  in the names list
 */
function getMatchName(names, name) {
  let position = names.indexOf(name);
  return position === -1 ? null : position+1;
}

/**
 *  Waits for the full results table to be populated
 *    @param {Page} thiSPage - Puppeteer page object
 *  @returns {Promise} Resolves when results table is ready
 */

/*
  <table class="Results-table Results-table--compact js-ResultsTable">
    <thead>...</thead>
    <tbody class="js-ResultsTbody">
      <tr class="Results-table-row" ...
*/
async function waitForResults(
  thisPage,
  timeSecs = loadDetailSECS)
{
  // await thisPage.waitForSelector('.js-ResultsTbody .Results-table-row',
  //  { visible: true, timeout: timeSecs*1000});
  const selectAllRUNNERS = 'tr.Results-table-row';
  try {
    await thisPage.waitForFunction((selectAllRUNNERS) => {
      return document.querySelectorAll(selectAllRUNNERS).length > 1;
    }, {timeout: timeSecs*1000},selectAllRUNNERS);
  } catch(err) {
    console.warn('WARNING: Unexpected delay retrieving detailed results on '+thisPage.url()+
      ' - consider either retrying later or checking whether domain Cookies have expired');
    throw new Error('ERROR: Failed to load detailed results '+err);
  }
}
  
/**
 * Orders the runners according to Age-Grade %age or reverts to default order of finish
 *    @param {Page} thisPage - Puppeteer page object containing table of runners
 *    @param {selection} order - typically agegrade-desc (unless to reset as default)
 */
/*
  CAUTION: Text may differ in different languages; sort Age-Grade on last option (-1)
  <select name="sort" class="js-ResultsSelect">
    <option value="position-asc">Sort by Position ▲</option>
    <option value="position-desc">Sort by Position ▼</option>
    <option value="runs-asc">Sort by Total parkruns ▲</option>
    <option value="runs-desc">Sort by Total parkruns ▼</option>
    <option value="agegrade-asc">Sort by Age Grade % ▲</option>
    <option value="agegrade-desc">Sort by Age Grade % ▼</option>
  </select>
*/
async function sortPositions(
  thisPage,
  order = 'position-asc')    // top option is default, overall position from 1..n (lowest time first)
{    // expect same dataset that may be quickly re-ordered getting Age-Grade positions prior to others
  const sortSelector = 'select[name="sort"]';
  const selectSortedRUNNERS = 'tr.Results-table-row';
  await thisPage.evaluate((order,sortSelector) => {
    let sortSelect = document.querySelector(sortSelector);    // valid inside evaluate
    console.log('sortPositions sortSelect:',sortSelect);
    sortSelect.value = order;
    sortSelect.dispatchEvent(new Event('change',{bubbles: true}));
  },order,sortSelector);
  await thisPage.waitForFunction((selectSortedRUNNERS) => {
    return document.querySelectorAll(selectSortedRUNNERS).length > 0;
  }, {timeout: loadDetailSECS*1000},selectSortedRUNNERS);
}

/**
 *  Effects the Age-Grade (descending) order in the runner results to match the position
 *    @param {Page} thisPage - Puppeteer page object containing table of results
 *    @param {string} matchnumber - name to match on this page of results 
 *  @returns {array} [position,numRunners] with ordered position within total number of runners
 *    ...and resets the order back to normal finishing-time (fastest first)
 */
async function sortAgeGrade(thisPage,matchRunner,ageGrade) {
  try {
    await waitForResults(thisPage);
    const numRunners = (await getRunnerNames(thisPage)).length;  // used as a pre-requisite for filter (later) as well as for sort
    console.log('Total number of runners is '+numRunners+' in results, '+thisPage.url());
    await sortPositions(thisPage,'agegrade-desc');
    let sortedRunners = await getRunnerNames(thisPage);
    // always expect fewer because some unknown or have not specified age/gender
    let numSortedRunners = sortedRunners.length;
    if (!numSortedRunners)
      throw new Error('Failed to find any runners by '+ageGrade);
    else if (numSortedRunners === numRunners)
      console.warn('WARNING: Consider no Unknown runners when sorted by '+ageGrade+' within results, '+thisPage.url());
    let position = getMatchName(sortedRunners,matchRunner);
    if (!position)
      throw new Error('Failed to match runner, '+matchRunner+' sorted by '+ageGrade+' within results, '+thisPage.url());
    await thisPage.reload();    // reset for next filter
    console.log(ageGrade+' position for matching runner, '+matchRunner+' is '+position+' (out of '+numSortedRunners+'/'+numRunners+')');
    return [position,numRunners];
  } catch (err) {
    console.error(err+' on '+thisPage.url());
    throw err;
  }
}

/*
// BEFORE entering the age Category, there is no class="item ..."
  <input type="text" name="search" class="js-ResultsSearch selectized"
    placeholder="Start typing to search" tabindex="-1" value="" style="display: none;">
  <div class="selectize-control js-ResultsSearch multi plugin-remove_button">
    <div class="selectize-input items not-full has-options">
      <input type="text" autocomplete="off" tabindex="" 
        placeholder="Start typing to search" style="width: 153px; opacity: 1; position: relative; left: 0px;">
    </div>
        :
        // sample subset category for Gender, Age Group, & Achievement
        <div class="option" data-selectable="" data-value="gender: Male">
          <span class="value">Male</span>
          <span class="type type--gender">Gender</span>
        </div>
        // samples after typing VM, which automatically gets highlighted
        <div class="option" data-selectable="" data-value="agegroup: VM35-39">
          <span class="value"><span class="highlight">VM</span>35-39</span>
          <span class="type type--agegroup">Age Group</span>
        </div>
// AFTER entering the age Category, the value is evident in both the visible and the hidden fields
  <input type="text" name="search" class="js-ResultsSearch selectized"
    placeholder="Start typing to search" tabindex="-1" value="agegroup: VM55-59" style="display: none;">
  <div class="selectize-control js-ResultsSearch multi plugin-remove_button">
    <div class="selectize-input items not-full has-options">
      <div class="item item--agegroup active" data-value="agegroup: VM55-59">
        <span class="filter-type ">Age Group: </span><span class="filter-value">VM55-59</span>
        <a href="javascript:void(0)" class="remove" tabindex="-1" title="Remove">×</a>

// BEFORE On Japanese events, age-category is similar to English but the gender category is Japanese chars...
  <input type="text" name="search" class="js-ResultsSearch selectized"
    placeholder="検索キーワードを入力してください。" tabindex="-1" value="" style="display: none;">
  <div class="selectize-control js-ResultsSearch multi plugin-remove_button">
    <div class="selectize-input items not-full has-options">
      <input type="text" autocomplete="off" tabindex="" style="width: 233px; opacity: 1; position: relative; left: 0px;"
        placeholder="検索キーワードを入力してください。">
    </div>
    <div class="selectize-dropdown multi js-ResultsSearch plugin-remove_button" style="display: none; width: 586px; top: 50px; left: 0px; visibility: visible;">
      <div class="selectize-dropdown-content">
        <!-- options will be listed here -->
      </div>
    </div>
  </div>

// AFTER pressing Enter in Japanese
  <input type="text" name="search" class="js-ResultsSearch selectized"
    placeholder="検索キーワードを入力してください。" tabindex="-1" value="gender: 男子" style="display: none;">
  <div class="selectize-control js-ResultsSearch multi plugin-remove_button">
    <div class="selectize-input items not-full has-options has-items">
      <div class="item item--gender active" data-value="gender: 男子">
        <span class="filter-type ">性別: </span>
        <span class="filter-value">男子</span>
        <a href="javascript:void(0)" class="remove" tabindex="-1" title="Remove">×</a>
// AFTER pressing Enter in Finnish, there is a minor difference in that the active state does not appear?
  <input type="text" name="search" class="js-ResultsSearch selectized"
    placeholder="Kirjoita hakutiedot" tabindex="-1" value="gender: Miehet" style="display: none;">
  <div class="selectize-control js-ResultsSearch multi plugin-remove_button">
    <div class="selectize-input items not-full has-options has-items">
      <div class="item item--gender" data-value="gender: Miehet">
        <span class="filter-type ">Sukupuoli: </span>
        <span class="filter-value">Miehet</span>
        <a href="javascript:void(0)" class="remove" tabindex="-1" title="Remove">×</a>

// Alternatively, select the pull-down option without typing!
  <div class="selectize-dropdown multi js-ResultsSearch plugin-remove_button"
      style="display: none; width: 485px; top: 50px; left: 0px; visibility: visible;">
    <div class="selectize-dropdown-content">
      :
      <div class="option" data-selectable="" data-value="achievement: 初参加!">
        <span class="value">初参加!</span><span class="type type--achievement">達成</span></div>
      <div class="option" data-selectable="" data-value="gender: 男子">
        <span class="value">男子</span><span class="type type--gender">性別</span></div>
      :
      <div class="option" data-selectable="" data-value="gender: 女子">
        <span class="value">女子</span><span class="type type--gender">性別</span></div>
      :
      <div class="option" data-selectable="" data-value="agegroup: JM11-14">
        <span class="value">JM11-14</span><span class="type type--agegroup">年齢層</span></div>

*/
/**
 * Retrieves a position after filtering runners according to Age or Gender Category and  reverts to default order
 *    @param {Page} thisPage - Puppeteer page object containing table of runners
 *    @param {number} numRunners - total number of runners including Unknowns (or unspecified category)
 *    @param {string} category - Category to filter (specfic Age-group or Gender) 
 *    @param {string} catClass - Category value to check - e.g. JM10...VW70-74, Male/Female
 */
async function filterPositions(
  thisPage,numRunners,
  category,catClass = 'agegroup')                         // alternatively 'gender'
{
  const selectBASE = '.selectize-input';   
  const selectINPUT = selectBASE+' input';         // Finds (2nd) input text field (within the class element)
  const selectOPTIONS = '.selectize-dropdown-content';
  const expectedVALUE = catClass+': '+category;
  const selectCatOPTION = selectOPTIONS+' .option[data-value="'+expectedVALUE+'"]';
  const selectFilteredRUNNERS = 'tr.Results-table-row';
  console.log('Expected value:'+expectedVALUE);
  console.log('Selector:'+selectCatOPTION);
  try {
    await thisPage.waitForSelector(selectINPUT);
    await thisPage.click(selectINPUT);                   //  1.  Pre-requisite for selecting drop-down?
    await thisPage.type(selectINPUT," ");                //      ...to ensure options become visible
    await thisPage.waitForSelector(selectOPTIONS,        //  2.  Skip to the multi drop-down list of options
      {visible: true, timeout: loadDetailSECS*1000});                              //      ...that are visible
    let optionExists = await thisPage.evaluate((selectCatOPTION,expectedVALUE) =>
    {
      let option = document.querySelector(selectCatOPTION);
      if (option) {
        option.click();                                  //  3.  Select specific category option (not the first!)
        return true;
      } else {
        console.warn('WARNING: No runners matching '+expectedVALUE+' - consider correcting runner DoB or gender translation?');
        return false;
      }
    },selectCatOPTION,expectedVALUE);
    if (optionExists) {
      let filteredOk = await thisPage.waitForFunction((numRunners,selectFilteredRUNNERS) =>
      {                                                //  4.  Verify by checking reduced no. of runners in table  
        return document.querySelectorAll(selectFilteredRUNNERS).length < numRunners;
      },{timeout: loadDetailSECS*1000},numRunners,selectFilteredRUNNERS);
      if (filteredOk)
        console.log('Successfully filtered '+expectedVALUE);
    }
  } catch (err) {
    console.error('ERROR: Unable to click on option with data-value as'+expectedVALUE+'\n'+err);
  }
}

/**
*  This removes any category filter to complement the above filterPositions within filterCategory
*    @param {Page} page - Puppeteer page object
*    @param {string} category - Category to remove (for info only)
*  @returns {Promise} Resolves when filter removed
*/
/*
  <div class="selectize-control js-ResultsSearch multi plugin-remove_button">
  	<div class="selectize-input items not-full has-options has-items">
      <div class="item item--agegroup" data-value="agegroup: VM55-59">
        <span class="filter-type ">Age Group:</span>
        <span class="filter-value">VM55-59</span>
        <a href="javascript:void(0)" class="remove" tabindex="-1" title="Remove">×</a>
        <a href="javascript:void(0)" class="remove" tIs waitabindex="-1" title="Remove">×</a>
*/
async function removeFilter(thisPage,category) {    // REDUNDANT since replaced with thisPage.reload
  const expectedValue = '';  //    ...likewise with gender: Male/Female
  // This only resets the filter search, without impacting the sort order
  const selectBASE = '.selectize-input';    //same as for filterPositions
  const selectOUTPUT = selectBASE+' .item';
  const selectREMOVE = selectOUTPUT+' .remove';
  try {
    await thisPage.waitForSelector(selectOUTPUT);
    await thisPage.waitForSelector(selectREMOVE);
    await thisPage.evaluate((selectREMOVE) => {
      document.querySelectorAll(selectREMOVE).forEach(btn => btn.click());
      // assume table update is instant, if expected data already queried on client browser
    }, selectREMOVE);
    // console.log(elementHTML);
    await thisPage.waitForFunction((selectBASE) => {    // verify No has-items
      var input = document.querySelector(selectBASE);
      return !input || !input.classList.contains('has-items');
    },selectBASE);
    console.log('Filter for category, '+category+' removed');
  } catch (err) {
    console.warn('WARNING: No filter for category, '+category+' to remove: '+err);
  }
}

/**
 *  Effects a filter of the results by category (agegroup/gender) and then remopves that filter
 *    @param {Page} thisPage - Puppeteer page object containing results for an event
 *    @param {string} matchRunner - Name & Surname to find in filtered table of runners
 *    @param {number} numRunners - total number of runners including Unknowns (or unspecified category)
 *    @param {string} category - Category to filter (specfic Age-group or Gender) 
 *  @returns {number} category position via {Promise}
 */
async function filterCategory(
  thisPage,matchRunner,numRunners,
  category,
  catClass = 'agegroup')  //   // alternatively 'gender'
{
  // Assumes default order of run-time position is preset on runner list (position-desc)
  try {
    await waitForResults(thisPage);  // filter options useless without the data
    await filterPositions(thisPage,numRunners,category,catClass);
    let filteredRunners = await getRunnerNames(thisPage);
    // always expect fewer because some unknown or have not specified age/gender
    let numFilteredRunners = filteredRunners.length;
    if (!numFilteredRunners)
      throw new Error('Failed to filter any '+category+' runners');
    else if (numFilteredRunners === numRunners)
      throw new Error('Failed to filter ['+numFilteredRunners+'] runners by '+catClass);
    let position = getMatchName(filteredRunners,matchRunner);
    if (!position)
      throw new Error('Failed to match runner, '+matchRunner+' filtered by '+category+' within results, '+thisPage.url());
    await thisPage.reload();    // reset for next filter
    console.log(category+' position for matching runner, '+matchRunner+' is '+position+' (out of '+numFilteredRunners+'/'+numRunners+')'); 
    return position;
  } catch (err) {
    console.error(err+' on '+thisPage.url());
    throw err;
  }
}

/**
* Called by GAS UrlFetch function, (via browser function to switch below)
*   Example https://<GC service>.run.app?url=https://www.parkrun.org.uk/havant/results/638&rn=Dave%20BUSH&ac=VM55-59&gc=Male
* Returns 3 positions (in JSON format): Age & Gender Category filter after getting the Age-Grade (%age) order
*/
exports.filterUrl = async (req,res) => {
  // Default parameters in case no ? and & parameters passed
  let thisUrl = req.query?.url     || 'https://www.parkrun.org.uk/havant/results/638/';  // Sample parkrun event
  let matchRunner = decodeURIComponent(req.query?.rn || 'Dave BUSH');  // Sample runner at Havant parkrun #638
  let ageCat = req.query?.ac       || 'VM55-59';      // Age-Category filter for matching Dave (expect 2)
  let ageGrade = req.query?.ag     || 'Age-Grade';    // Age-Grade (%age) &sort for matching Dave (expect 8)
  let genderCat = req.query?.gc    || 'Male';         // Gender category filter for matching Dave (expect 11)
  let caching = Boolean(req.query?.cache === 'true');     // No caching in catch-up mode on a single runner at different events
// begin
  console.log('thisUrl: '+thisUrl);
  console.log('matchRunner: '+matchRunner);
  console.log('ageCat: '+ageCat);
  console.log('gcCat: '+genderCat);      // WARNING: Values differ per language/country
  console.log('ageGrade: '+ageGrade);
  console.log('caching: '+caching);
  // console.log('caching type: '+typeof caching);
  var thisPage;
  // TODO: Instead of caching multiple pages per import, need to consider session-based multiple pages per spreadsheet!
  // if (thisUrl in cachedPages) {        // typically, many runners at the same event (during weekly import only)
    // console.log('No. of cached pages since caching: '+Object.keys(cachedPages).length);
    // console.log('Re-using detailed cached results for URL, '+thisUrl);
    // thisPage = cachedPages[thisUrl];    // ...and so no delay in loading OR in awaiting enforced delay between each
  // } else {
    // cachedPages[thisUrl] = thisPage;
    // console.log('Cached results page for URL, '+thisUrl);
    // console.log('No. of cached pages after caching: '+Object.keys(cachedPages).length);
  // }
  try {
    if (caching && prevPage && thisUrl == prevFilterUrl) {
      thisPage = prevPage;
      console.log('Filter positions using same page for runner, '+matchRunner+' since common event URL, '+thisUrl);
    } else {
      thisPage = await loadUrl(thisUrl,-1,loadDetailSECS,caching);
      prevFilterUrl = thisUrl;
      prevPage = thisPage;
      var testCmd = 'curl -X GET "'+browserURL+'/filterUrl'+'?url='+thisUrl+'&rn='+matchRunner+'&ac='+ageCat+'&gc='+genderCat+'&cache='+caching+'" \\'
        +'-H "Authorization: bearer $(gcloud auth print-identity-token)" \\'
        +'-H "Content-Type: application/json"';
      console.log('Test: '+testCmd);
    }
  } catch (err) {
    console.error(err);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end('ERROR: Failed to load URL, '+thisUrl+' while caching is '+caching+' - '+ err);
    // res.status(500).send('ERROR: Failed to load URL, '+thisUrl+' while caching is '+caching ' + err);
  }
  if (thisPage) {
    try {  // Get 2 (or more) positions in series?
      // 1. Sort by (descending) Age-Grade, to get ageGrade position of matchRunner
      let [agPosition,numRunners] = await sortAgeGrade(thisPage,matchRunner,ageGrade);
      // 2. Filter by Age-Category to get ageCat position of matchRunner
      let acPosition = await filterCategory(thisPage,matchRunner,numRunners,ageCat);
      // 3. Filter by Gender if need to get genderCat position of matchRunner
      let gcPosition = await filterCategory(thisPage,matchRunner,numRunners,genderCat,'gender');
      res.writeHead(200, { 'Content-Type': 'text/plain' });
      // res.status(200).json({acPosition,agPosition});      // in expected order
      res.end(JSON.stringify({acPosition,agPosition,gcPosition}));    // in expected order
    } catch (err) {
      console.error('ERROR:',err);
      res.writeHead(500, { 'Content-Type': 'text/plain' });
      res.end('ERROR: '+err.message);
      // res.status(500).send('ERROR: '+err.message);
    } finally {
      // TODO: await thisPage.close();  // re-use page may fail??, consider new Page for each parkrun results instance
    }
  }
}

const cookieJAR = [
  'https://www.parkrun.org.uk',  // may be used interationally for any runner
  'https://www.parkrun.com',     // this is required for consolidated results for a group date
  'https://www.parkrun.co.nl',   // this and others may be used for an event results, wherever
  'https://www.parkrun.com.de',
  'https://www.parkrun.com.au',
  'https://www.parkrun.ca',
  'https://www.parkrun.jp',
  'https://www.parkrun.co.za'
];

async function deleteOldCookies(page,targetUrl) {
  try {
    let cookies = await page.cookies();
    if (await cookies.find(c => c.name === 'psc')) {
      await page.deleteCookie({name:'psc',url:targetUrl});
      await page.reload();
    }
  } catch (err) {
    console.log('Cookie for url,',targetUrl,'to be deleted was not found:',err);
  }
}

exports.deleteCookies = async (_,res) => {
  // No request because handles ALL parkrun domain URLs
  // TODO: This assumes initBrowser previously launched, and therefore ought to rely on the WSEP already set
  var thisBrowser;
  try {
    thisBrowser = await puppeteer.launch({  // variable delay if image not cached?
      headless: true,
      executablePath: chromePATH,
      userDataDir: myUserDataDir, // clear cookies and certs?
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--cert='+allParkrunCERTS,
        '--verbose',
      ],
      timeout: launchSECS*1000,       // max launch time
    });
  } catch (err) {
    let resultErr = 'ERROR: Failed to launch browser to delete Cookies\n'+err;
    console.error(resultErr);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end(resultErr +' - '+ err);
    // res.status(500).send(resultErr);
  }
  let thisPage = await thisBrowser.newPage();  // Re-use the same Page since not in parallel
  for (var domainUrl of cookieJAR) {
    try {
      await thisPage.goto(domainUrl,{waitUntil: 'domcontentloaded',timeout: 10000});
      let cookies = await thisPage.cookies(domainUrl);
      await thisPage.deleteCookie({name: 'psc', url: domainUrl});
    } catch (err) {
      console.error('ERROR: unable to delete parkrun cookie for '+domainUrl+'\n'+err);
    } finally {
      await thisPage.close();
      if (thisBrowser) await thisBrowser.close();
    }
  }
}

exports.acceptCookies = async (_,res) => {
  // No request because handles ALL parkrun domain URLs
  // TODO: This assumes initBrowser previously launched, and therefore ought to rely on the WSEP already set
  var thisBrowser;
  try {
    thisBrowser = await puppeteer.launch({  // variable delay if image not cached?
      headless: true,
      executablePath: chromePATH,
      userDataDir: myUserDataDir, // store cookies and certs
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--cert='+allParkrunCERTS,
        '--verbose',
      ],
      timeout: launchSECS*1000,       // max launch time
    });
  } catch (err) {
    let resultErr = 'ERROR: Failed to launch browser to check Cookies ok\n'+err;
    console.error(resultErr);
    res.writeHead(500, { 'Content-Type': 'text/plain' });
    res.end(resultErr +' - '+ err);
    // res.status(500).send(resultErr);
  }
  // CAUTION: Although any single parkrun certificate may suffice, ...
  //  ... do not assume the Cookie has been accepted on EACH and every domain!
  let thisPage = await thisBrowser.newPage();  // Potentially re-use the same Page since not in parallel
  let results = [];
  for (var domainUrl of cookieJAR) {
    try {
      await thisPage.goto(domainUrl,{waitUntil: 'domcontentloaded',timeout: 10000});
      const acceptButton = `button.cm__btn[data-role="all"]`;
      try {
        await thisPage.waitForSelector(acceptButton, {timeout: 5000});
        let accept = await thisPage.$(acceptButton);
        await thisPage.setCookie({
          name: 'psc',
          value: 'some-value',
          domain: domainUrl
        });
        var result;
        if (accept) {
          await new Promise(r => setTimeout(r, 2000)); // Wait 2 second for any animations to finish
          // await accept.click();
          await thisPage.click(acceptButton);
          result = 'Cookies accepted for site, '+domainUrl;
        } else   // otherwise assume already accepted and simply confirm so!
          result = 'Assume cookies already accepted for site, ' + domainUrl;
        console.log(result);
        results.push(result);
      } catch (warning) {    // If no Accept button appears, then that is the norm
        // WARNING Perhaps retry in case page not fully evaluated or delete and redo
        // TODO: Check logs in case failed for button not detected
        // deleteCookies(thisPage,domainUrl);
        let resultWarn = 'No button presented for Cookies to be accepted on site, '+domainUrl+'\n  WARNING: '+warning;
        console.warn(resultWarn);
        results.push(resultWarn);
      };
    } catch (err) {
      let resultErr = 'ERROR: Failed to access parkrun domain, '+domainUrl+' to check whether cookie accepted:\n  '+err;
      console.error(resultErr);
      results.push(resultErr);
    };
  }
  if (thisPage) await thisPage.close();
  await thisBrowser.close();
  res.writeHead(200, { 'Content-Type': 'text/plain' });
  res.end(results.join('<br><br>').replace(/\n/g, '<br>'));
  // res.status(200).send(results.join('<br><br>').replace(/\n/g, '<br>'));
}

/**
*  This browser function provides a convenient single entry point (as defined in package.json).
*  Nevertheless, the delegated URL-based functions are equally valid entry and exit points.
"  Therefore, await is redundant in calling these functions because each effect the return directly.
*
*  INFO : The return parameter, 'res' is critical within the (Node.js) delegated functions:
*    1. .status(200) to return the HTTP statusCode (where 200 = success)
*    2. .type() to set the Content-Type header
*          - normally, implcit from content: text/plain (default), text/html or application/json
*          - alternatively, explicitly also using res.setHeader('Content-Type', type)
*    3. .send(body) to return the body AND to end the call back to the Google Spreadsheet client
*          - alternativly, .end may return plain/text or follow (one or more) .write
*/
exports.browser = async (req,res) => {
  var parsedUrl = url.parse(req.url,true);
  req.query = parsedUrl.query;
  var path = parsedUrl.pathname;
  console.log('Path identified: ',path);
  if (path === '/initBrowser')
    exports.initBrowser(req,res);
  else if (path === '/getUrl')
    exports.getUrl(req,res);
  else if (path === '/filterUrl')
    exports.filterUrl(req,res);
  else if (path === '/stopBrowser')
    exports.stopBrowser(req,res);
  else if (path === '/acceptCookies')
    exports.acceptCookies(req,res);
  else if (path === '/deleteCookies')
    exports.deleteCookies(req,res);
  else if (path.includes('wp-content'))   // ignore internal relative dependencies
    console.log('INFO: Ignoring local dependencies: '+path);
  else if (path === '/favicon.ico') {     // ...including icons
    res.writeHead(204, { 'Content-Type': 'image/x-icon' });
    res.end();
  } else {
    console.log('ERROR: Invalid Puppet function path,',path);
    res.writeHead(404, { 'Content-Type': 'text/plain' });
    res.end('ERROR: Invalid Puppet function path, '+path);
  } 
}

// Define the server simply
const nodeServer = http.createServer((req, res) => {
    // Direct handoff to your existing logic
    exports.browser(req, res).catch(err => {
        console.error("ERROR: ", err);
    });
});

// Start listening on port 36007
nodeServer.listen(PORT, HOST, () => {
  console.log('=====================================================================================');
  console.log('🚀 Local Node Server is active on '+browserURL);
  console.log('=====================================================================================');
});
