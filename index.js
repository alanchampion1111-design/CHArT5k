// Multiple steps to do:
//    0. Verify base image is in Artifact Registry (done)
//    1. Verify triggers gets latest sources from GitHub including this index.js file (done)
//    2. Verify build image uses Docker to install Chrome (done)
//    3. Verify Chrome browser works directly on server after build (done)
//    4. Verify server side, thisBrowser activated by client from Google Spreadsheet app (done)
//    5. Verify sample page retrieved ok and that thisBrowser and thisPage persists (done)
//    6. Upload profile/certificates for access to www.parkrun.org.uk (done)
//    7. Verify allowed to load content for www.parkrun.org.uk (done)
//    8. Verify stealth access to individual parkrunner results table (tbd - although disallowed)

// const functions = require('@google-cloud/functions-framework');
// const puppeteer = require('puppeteer');
const puppeteer = require('puppeteer-extra');
const StealthPlugin = require('puppeteer-extra-plugin-stealth');
puppeteer.use(StealthPlugin());
const userAgent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36';
const url = require('url');
let thisBrowserWSEp;  // browser persists on server   
const browserURL = 'https://browser-automation-service-224251628103.europe-west1.run.app';
const parkrunURL = 'https://www.parkrun.org.uk';
const parkrunnerURL = parkrunURL+'/parkrunner/';

let thisPageId;       // re-use same page      
let browserTimeout;   // for browser session
let browserTimer;
const launchSECS = 45000;
const pageSECS = 10000;   // minimum of 10 seconds between page accesses on parkrun site
let initPromise;      // browser "finished" after initialised (although still active

/**
 *  Launches a headless Chrome browser with specified session limit.
 *    @param {number} [sessionLimit default 5] - Session limit in minutes
 *  @returns {Promise<void>}
 *  @sideeffect leaves the browser connected and returns a presistent WS endpoint for re-use
 */
let cloudBrowser = async (
  sessionLimit = 55) =>
{
  browserTimeout = sessionLimit*60*1000;
  var thisBrowser = await puppeteer.launch({  // variable delay if image not cached?
    headless: true,
    executablePath: '/usr/bin/google-chrome',
    args: [
      '--no-sandbox',
      '--disable-setuid-sandbox',
      '--disable-dev-shm-usage',
      '--cert=./www.parkrun.org.uk.pem',  // Certificate running on UK site, independent of country of event?
      '--verbose',
      '--lang=en-GB'    //  ensures the date formats appear as dd/mm/yyyy
    ],
    timeout: launchSECS,       // max launch time
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
      initPromise = undefined;
      thisBrowserWSEp = null;
      thisPageId = null;
      clearTimeout(browserTimer);
    }
  }, browserTimeout);
  thisBrowserWSEp = thisBrowser ? thisBrowser.wsEndpoint() : null;  // return to client (although also global)
  console.log('Retained browser WS Endpoint:',thisBrowserWSEp);
  try {
    var thisPage = await thisBrowser.newPage();
    if (thisPage) {
      thisPageId = await thisPage.target()._targetId;
      console.log('Retained page ID,',thisPageId);
      thisPage.setDefaultTimeout(pageSECS);  // Set the timeout for loading the page
      await thisPage.setUserAgent(userAgent);
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


/**
 *  Initializes the browser and returns the WebSocket endpoint.
 *    @returns {Promise<void>} before continuing
 *    @sideeffect - preserve the browser WS endpoint and re0usable page ID for re-use
 */
exports.initBrowser = async (_,res) => {
  if (!initPromise) {
    initPromise = (async () => {
      try {
        await cloudBrowser(60);  // Launched ok, but browser active in background
        res.status(200).send(thisBrowserWSEp);
      } catch (err) {
        console.error('ERROR: Failed to initialise browser:',err);
        // consider a relaunch with args, --pull from Docker if image is not properly cached!
        res.status(500).send('ERROR: Failed to initialise browser, '+err);
      } finally {
        // NEVER disconnect because this loses the puppeteer Stealth (plugin) setting!
        // await thisBrowser.disconnect();
        console.log('Returning immediately after (attempt at) launching browser');
      }
    })();
  } else {  // do nothing because browser previously launched
    res.status(200).send(thisBrowserWSEp);
  }
}

/**
 *  Loads URL in Puppeteer and waits for page load
 *    @param {Page} page - Puppeteer page object
 *    @param {string} url - URL to load
 *  @returns thisPage for detailed searchning or the HTML content via a {Promise} Resolves when page loaded
 */
let loadUrl = async (thisUrl, pageOnly=false) => {
  // console.log('Reconnecting to browser WS Endpoint:',thisBrowserWSEp,'with same page ID,',thisPageId);
  try {
    if (!thisBrowserWSEp) {
      console.error('ERROR: Persistent browser NOT found:',thisBrowserWSEp);
      throw new Error('Persistent browser NOT found!');
    } else {
      console.log('Persistent browser found,',thisBrowserWSEp,'with ongoing timeout:',browserTimeout);
      var thisBrowser = await puppeteer
        .connect({browserWSEndpoint: thisBrowserWSEp, timeout: browserTimeout});     // actually reconnects 
      var thisPage = (await thisBrowser.pages())
        .find(page => page.target()._targetId === thisPageId);
      if (thisPage) {
        await thisPage.setDefaultTimeout(pageSECS);
      } else {
        console.error('ERROR: Persistent page NOT found:',thisPageId,'with refreshed timeout', pageSECS);
        throw new Error('Persistent page NOT found!');
      }
      console.log('Persistent browser timeout,',browserTimeout,'with inter-page access delay,',pageSECS);
      console.log('Loading page with URL,',thisUrl);
      await thisPage.goto(thisUrl,{waitUntil: 'domcontentloaded'});
      // await thisPage.waitForFunction(() => window.parkrunResultsData);  // although not fully formaatted?
      var content = await thisPage.content();   // always ensure page is fully loaded
      return pageOnly ? thisPage : content;      // if content, then we are done, otherwise more to do!
    }
  } catch (err) {
    console.error('ERROR: Failed to retrieve page:',err);
    throw err;
  }
  // TODO: Consider closing the page and avoid re-using if parallel approach better and less risky
}

/**
 *  Gets the entire HTML content for a URL, typically all of results for a specific parkrunner 
 *    @param {string} url - e.g. 
 *  @returns {string} as HTML content
 */
exports.getUrl = async (req,res) => {
  // Default in case no ? parameters passed - sample runner is Alan
  let thisUrl = req.query?.url || 'https://www.parkrun.org.uk/parkrunner/777764/all/';
  try {
    var content = await loadUrl(thisUrl);
    res.status(200).send(content);
  } catch (err) {
    console.error(err);
    res.status(500).send('ERROR: Failed to load URL, '+thisUrl);
  } finally {
    // delay between calls (before any returns) while browser remains active
    await new Promise(resolve => setTimeout(resolve,pageSECS));
    // but NEVER disconnect because this loses the puppeteer Stealth (plugin) setting!
    // await thisBrowser.disconnect(); 
  }
}

/** ______________________________________________________________________________________
/
  Functions hirarchy follows for this block used maintain 
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
        resopmoveCategory
/ ________________________________________________________________________
*/

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
  await thisPage.waitForSelector(resultsTABLE);
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
 *  Waits for the 2nd sortable results table to be populated
 *    @param {Page} thiSPage - Puppeteer page object
 *  @returns {Promise} Resolves when results table is ready
 */
async function waitForResults(thisPage) {
  await thisPage.waitForFunction(() => {
    var tables = document.querySelectorAll('table.sortable');
    return tables.length >= 2;
  }, {timeout: 5000, polling: 1000});
}
  
/**
 * Orders the runners according to Age-Grade %age or reverts to default order of finish
 *    @param {Page} thisPage - Puppeteer page object containing table of runners
 *    @param {selection} order - typically agegrade-desc (unless to reset as default)
 */
/*
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
  await thisPage.evaluate((order) => {
    const sortField = 'sort';
    const sortSelector = `select[name="${sortField}"]`;
    let sortSelect = document.querySelector(sortSelector);    // valid inside evaluate
    console.log('sortPositions sortSelect:',sortSelect);
    sortSelect.value = order;
    sortSelect.dispatchEvent(new Event('change',{bubbles: true}));
  }, order);  // ensures order is in scope of the thisPage evaluation
}

/**
 *  Effects the Age-Grade (descending) order in the runner results to match the position
 *    @param {Page} thisPage - Puppeteer page object containing table of results
 *    @param {string} matchnumber - name to match on this page of results 
 *  @returns {number} position of matching name via {Promise} Resolves when sorted
 *    ...and resets the order back to normal finishing-time (fastest first)
 */
async function sortAgeGrade(thisPage,matchRunner,ageGrade) {
  try {
    await waitForResults(thisPage);  // sort options useless without the data
    await sortPositions(thisPage,'agegrade-desc');
    let runners = await getRunnerNames(thisPage);
    console.log('Number of '+ageGrade+' runners found: '+runners.length);
    if (!runners) throw new Error('Failed to find any runners by '+ageGrade);
    let position = getMatchName(runners,matchRunner);
    if (position) console.log(ageGrade+' position for matching runner, '+matchRunner+' is '+position);
    else throw new Error('Failed to find matching runner, '+matchRunner+' in sorted '+ageGrade+' within results, '+thisPage.url());
    await sortPositions(thisPage); // Reset to default order before getting next order
    return position;
  } catch (err) {
    console.error(err,'on',thisPage.url());
    throw err;
  }
}

/**
 * Retrieves a position after filtering runners according to Age or Gender Category and  reverts to default order
 *    @param {Page} thisPage - Puppeteer page object containing table of runners
 *    @param {selection} order - typically agegrade-desc (unless to reset as default)
 */
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
*/
async function filterPositions(
  thisPage,category,
  catClass = 'agegroup')                         // alternatively 'gender'
{
  const expectedValue = catClass+': '+category;  //    ...likewise with gender: Male/Female
  // var initialRowCount = await thisPage.locator('.table-selector tr').count();  // Check WARNING below?
  // const searchINPUT = 'input#search';         // FAILS -Does not find 1st input field since hidden!!
  const selectBASE = '.selectize-input';
  const selectOUTPUT = selectBASE+' .item';     
  const selectINPUT = selectBASE+' input';       // Finds 2nd input text field (within the class element)
  await thisPage.waitForSelector(selectINPUT);
  await thisPage.click(selectINPUT);             //  1. Focus may be automatic on typing in 2.
  await thisPage.type(selectINPUT,category);     //  2. Type valid Age-Category (or Male/Female Gender)
  await thisPage.keyboard.press('Enter');        //  3. Press Enter to select matching pull-down...
  // var elem = await thisPage.$(selectBASE);       
  // console.log(await elem.evaluate( elem => elem.outerHTML));  // ...confirmed as expected in commit b248b74
  await thisPage.waitForSelector(selectOUTPUT,   //  4. Wait until the new element exists
    {visible: true,timeout: 5000});              //     ... with table mods also expected?
  let selectedValue = await thisPage.$eval(      //  5. Verify match to pull-down in the item that follows...            
    selectOUTPUT,elem => elem.dataset.value);    //      ...as likewise directed into searchINPUT (but hidden!)
  if (selectedValue === expectedValue)
    console.log('The filter option for '+category+' matched a pull-down option');
  else {
    elem = await thisPage.$(selectBASE);
    console.log(await elem.evaluate(elem => elem.outerHTML));
    throw new Error('Expected '+expectedValue+' category but got '+selectedValue);
  }
  // TODO: Perhaps may also await the table update, but may be handled without re-query on the browser side?
  // Assume table update of row subset is instant? if filter handled locally by scripts
  // WARNING: If this fails because of a backend service, consider continue only after number of rows differ
  // await thisPage.waitForFunction((initialRowCount) => {
  //   var newRowCount = document.querySelectorAll('.table-selector tr').length;
  //   return newRowCount !== initialRowCount;
  // }, initialRowCount);
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
        <a href="javascript:void(0)" class="remove" tabindex="-1" title="Remove">×</a>
*/
async function removeFilter(thisPage,category) {
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
 *    @param {string} category - Category to filter (Age-group or Gender) 
 *  @returns {number} category position via {Promise}
 */
async function filterCategory(
  thisPage,matchRunner,
  category,
  catClass= 'agegroup')   // alternatively 'gender' 
{
  // Assumes default order of run-time position is preset on runner list (position-desc)
  try {
    await waitForResults(thisPage);  // filter options useless without the data
    await filterPositions(thisPage,category,catClass);
    let runners = await getRunnerNames(thisPage);
    console.log('Number of '+category+' runners found: '+runners.length);
    if (!runners.length) throw new Error('Failed to filter on '+category+' category');
    let position = getMatchName(runners,matchRunner);
    if (position) console.log(category+' position for matching runner, '+matchRunner+' is '+position);
    else throw new Error('Failed to find matching runner, '+matchRunner+' in filtered '+category+' within results, '+thisPage.url());
    await removeFilter(thisPage,category); // Reset filter WHEN a subsequent position is required (e.g. gender)
    return position;
  } catch (err) {
    console.error(err, 'on', thisPage.url());
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
// begin
  console.log('thisUrl: '+thisUrl);
  console.log('matchRunner: '+matchRunner);
  console.log('ageCat: '+ageCat);
  console.log('gcCat: '+genderCat);
  console.log('ageGrade: '+ageGrade);
  var testCmd = 'curl -X GET "'+browserURL+'/filterUrl'+'?url='+thisUrl+'&rn='+matchRunner+'&ac='+ageCat+'&gc='+genderCat+'"" \\'
    +'-H "Authorization: bearer $(gcloud auth print-identity-token)" \\'
    +'-H "Content-Type: application/json"';
  console.log('Test: '+testCmd);
  var thisPage = await loadUrl(thisUrl,true);
  try {  // Get 2 (or more) positions in series
    // 1. Sort by (descending) Age-Grade, to get ageGrade position of matchRunner
    let agPosition = await sortAgeGrade(thisPage,matchRunner,ageGrade);
    // 2. Filter by Age-Category to get ageCat position of matchRunner
    let acPosition = await filterCategory(thisPage,matchRunner,ageCat);
    // 3. Filter by Gender if need to get genderCat position of matchRunner
    let gcPosition = await filterCategory(thisPage,matchRunner,genderCat,'gender');
    // res.status(200).json({acPosition,agPosition});      // in expected order
    res.status(200).json({acPosition,agPosition,gcPosition});    // in expected order
  } catch (err) {
    console.error('ERROR:',err);
    res.status(500).send('ERROR: '+err.message);
  } finally {
    // TODO: await thisPage.close();  // re-use page may fail??, consider new Page for each parkrun results instance
    console.warn('WARNING: If re-using the same page, the normal parallel performance may be slower (or otherwise interfere)');
  }
}

exports.stopBrowser = async (_,res) => {
  try {
    if (thisBrowserWSEp) {
      var thisBrowser = await puppeteer
        .connect({ browserWSEndpoint: thisBrowserWSEp });
      if (thisPageId) {
        var thisPage = (await thisBrowser.pages())
          .find(page => page.target()._targetId === thisPageId);
        if (thisPage) {
          await thisPage.close();
          console.log('Page closed successfully - Page Id:',thisPageId);
        } else {
          console.warn('WARNING: Page previously closed or timed out - Page Id:',thisPageId);
        }
      }
      if (thisBrowser && thisBrowser.isConnected()) {
        await thisBrowser.close();
        console.log('Browser terminated successfully - WS endpoint:',thisBrowserWSEp);
        res.status(200).send('Browser terminated successfully');
      } else {
        console.warn('WARNING: Browser previously aborted or timed out - WS endpoint:',thisBrowserWSEp);
        res.status(204).send('WARNING: Browser previously aborted or timed out');
      }
    } else {
      console.warn('WARNING: Browser previously terminated - WS endpoint:',thisBrowserWSEp);
      res.status(204).send('WARNING: Browser previously terminated');
    }
  } catch (err) {
    console.error('ERROR: Failed to close page and/or terminate browser:',err);
    res.status(500).send('ERROR: Failed to close page and/or terminate browser, '+err);
  } finally {  // executed in all cases, even before the returns
    initPromise = undefined;
    thisBrowserWSEp = null;
    thisPageId = null;
    clearTimeout(browserTimer);
  }
}

async function deleteCookies(page,targetUrl) {
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

exports.acceptCookies = async (_,res) => {
  let cookieJar = [
    'https://www.parkrun.org.uk',
    'https://www.parkrun.com'
  ];
  
  try {
    let thisBrowser = await puppeteer.launch({  // variable delay if image not cached?
      headless: true,
      executablePath: '/usr/bin/google-chrome',
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--cert=./www.parkrun.org.uk.pem',
        '--verbose',
      ],
      timeout: launchSECS,       // max launch time
    });
    let thisPage = await thisBrowser.newPage();
    try {
      thisCookieURL = cookieJar[0];    // assume org.uk Cookie will suffice
      await thisPage.goto(thisCookieURL,{waitUntil: 'domcontentloaded',timeout: 10000});
      const acceptButton = `button.cm__btn[data-role="all"]`;
      try {
        thisPage.waitForSelector(acceptButton, {timeout: 10000});
        await thisPage.setCookie({
          name: 'psc',
          value: 'some-value',
          domain: 'www.parkrun.org.uk'
        });
        await thisPage.click(acceptButton);
        console.log('Cookies accepted for sites,',cookieJar);
        res.status(200).send('Required Cookies accepted for sites, '+cookieJar);
      } catch (warning) {    // If no Accept button appears, then that is the norm
        // WARNING Perhaps retry in case page not fully evaluated or delete and redo
        // TODO: deleteCookies(thisPage,thisCookieURL);
        console.warn('WARNING:',warning);    // check logs in case failed for button not detected
        console.log('No button presented for Cookies to be accepted on sites, ',cookieJar);
        res.status(200).send('No button presented for Cookies to be accepted on sites, '+cookieJar);
      }
    } catch (err) {
      console.error('ERROR: Failed to load page,',thisCookieURL,' to check Cookies:',err);
      res.status(500).send('ERROR: Failed to load page, '+thisCookieURL+' to check Cookies ok: '+err);
    }
  } catch (err) {
    console.error('ERROR: Failed to launch browser to check Cookies ok:',err);
    res.status(500).send('ERROR: Failed to launch browser to check Cookies ok: '+err);
  } finally {
    if (thisBrowser) await thisBrowser.close();
  }
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
  var path = parsedUrl.pathname;
  if (path === '/initBrowser') {
    exports.initBrowser(req,res);
  } else if (path === '/getUrl') {
    exports.getUrl(req,res);
  } else if (path === '/filterUrl') {
    exports.filterUrl(req,res);
  } else if (path === '/stopBrowser') {
    exports.stopBrowser(req,res);
  } else if (path === '/acceptCookies') {
    exports.acceptCookies(req,res);
  } else {
    console.log('ERROR: Invalid Cloud Run function path,',path);
    res.status(404).send('ERROR: Invalid Cloud Run function path, '+path);
  } 
}
