// FILE: AppCode.gs

/**
 * Global App Tables (Owner protected except Devices is open to all)
 */
const TABLES = {
  MASTER: "CHArT5k",      // contains list of Groups with summary extracts
  DEVICES: "Devices",     // ALL App users' settings locked to their device
  MEMBERS: "Results Gids" // composite of members' results across ALL Groups
  // A table also exists per named group; hence COL_GROUP below
};

/**
 * Master CHArT5k Table Headers (one row per registered Group)
 * Each Group Overview MUST be pulled by App Owner (triggered by Group owner?)
 *  to avoid blockages with dynamic accesses during App usage
 */
const COL_MASTER = {
  // DRIVER: "Active SS Id",  // pending IMPORTRANGE formula => triggered pull extracts
  READY: "Ready",             // initial results imported? (with About 1-liner in Note)
  GROUP_NAME: "Spreadsheet",  // configured Group spreadsheet (with About desc in Note )
  SS_ID: "SS Id",             // ...with its unique linkable Google worksheet id,
  OWNER: "Owner",             // ...eventual ownership transferred from founder
  RUNNERS_GID: "Runners Gid", // links to Group's Runners details sheet,
  LATEST_GID: "Latest Gid",   // ...latest results presented after weekly import,
  RANKINGS_GID: "Rankings Gid",     // e.g. current ranking used for League divisions
  CHALLENGES_GID: "Challenges Gid", // ordered results from planned event (same venue?)
  COUNT: "#",                       // Number of members imported
  GROUP_ID: "Parkrun Group No."     // optional (if it exists?)
  // Other virtual columns include generated latest/current/ranking table links
};

/**
 * Group Sheet Table Headers (for ALL comparative Chart Sheets in Group, in 4 sets of divisions, perhaps up to ~20 per set)
 */
const COL_GROUP = {
  CHART_NAME: "Performance Chart",  // name indicates division by Grade/Junior/Senior etc.
  SS_ID: "SS Id",                 // within unique linkable Google worksheet id,
  TARGET_GID: "Perf Gid",         // links to a single Group chart
  PERF_SHEET: "Perf Sheet",       // categorised under Age Groups, Leagues, etc. 
  CELL_RANGE: "Range",            // used to direct App user to a location on the sheet
  CHART_ID: "Chart Id",           // for highlighting App user line? (or for copying)
  COUNT: "#",                     // No. of runners listed on chart
  RUNNERS_IDS: "Runners Ids"      // list of Runner Ids within the chart division
  // Other virtual columns include the generated chart link
};

/**
 * Members Table Headers (lookup ALL members across ALL Groups)
 * Composite set MUST be pulled by App Owner (trigger)
 */
const COL_MEMBER = {
  RUNNER_ID: "Runner Id",     // A: Runner Id,only unique per Group (SS Id)
  SS_ID: "SS Id",             // B: composite set covers all Groups (SS Id 
  RESULTS_GID: "Results Gid", // C: automatic from Runner Id (as sheet name) in SS Id
  COUNT: "#",                 // D: No. of results per member
  SHARED: "Shared"            // E: Viewer/Commenter/Editor with Google-valid email address
};

/**
 * Devices Table Headers (subset of ALL App users' profile settings)
 *  Cache all field values for App user profile, only for App user)
 */
const COL_DEVICE = {          // row only exists per actual App user
  TIMESTAMP: "Timestamp",     // A: when user first used (introduced for Web App)
  DEVICE_ID: "Device Id",     // B: device locked to App user's settings
  GROUP_NAME: "Spreadsheet",  // C: user-selected Group (with extended guide in Note)
  SS_ID: "SS Id",             // D: automatic ref from selected Group (used for linking)
  RUNNER_ID: "Runner Id",     // E: user-selected Identity (with extended guide in Note)
  RESULTS_GID: "Results Gid", // F: automatic from Runner Id (as sheet name) in SS Id
  COUNT: "#",                 // G: No. of results per member (at time of setup)
  STATUS: "Access"            // H: Needed (default), Requested (share pending) or Granted (if Shared)
};

/**
 * Target Table/Chart Cell Location
 */
const VIEWPORTS = {
  TRENDS_OVERALL: "N2:P20",     // trend chart location fixed within each runner's results sheet
  TRENDS_RECENT: "N22:P40",     // trend chart below, focussed on past year (if one exists)
  RANKINGS_CURRENT: "A8:K109",   // current age-graded table in Rankings sheet location
  RANKINGS_BEST: "A111:K212",    // best-ever age-graded table (based on max  100 runners in Rankings)
};

/**
 * Web App Entry Point - Serves the HTML frontend shell
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('CHArT5k')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Core initialization called by the client on boot.
 * Verifies if the device UUID is already linked to a profile and delivers all guide parameters.
 */
function initializeApplicationData(devId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1A. Fetch 'About' info from MASTER table (via Header Notes)
    var masterSheet = ss.getSheetByName(TABLES.MASTER);
    var aboutShort = "Welcome to CHArT5k comparative charts for your parkrun groups";
    var aboutLong = "No matter how far apart you live, share your club or family group parkrun experience together much closer!";
    if (masterSheet) {
      var masterHeaders = masterSheet.getRange(1, 1, 1, masterSheet.getLastColumn()).getValues()[0];
      // Locate the "Ready" column to pull the short descriptive pitch
      var idxReadyCol = masterHeaders.indexOf(COL_MASTER.READY);
      if (idxReadyCol > -1) {
        var noteShort = masterSheet.getRange(1, idxReadyCol + 1).getNote();
        if (noteShort && noteShort.trim()) aboutShort = noteShort.trim();
      }
      // Locate the "Spreadsheet" column to pull the long narrative description
      var idxGroupCol = masterHeaders.indexOf(COL_MASTER.GROUP_NAME);
      if (idxGroupCol > -1) {
        var noteLong = masterSheet.getRange(1, idxGroupCol + 1).getNote();
        if (noteLong && noteLong.trim()) aboutLong = noteLong.trim();
      }
    }

    // 1B. Fetch 'User profile setting' from DEVICES table (via Header Notes)
    var deviceSheet = ss.getSheetByName(TABLES.DEVICES);
    var userSetupGuide = "Select your club or family group with your runner Id (if known)";
    var memberRequestGuide = "Select your club or family group (if it exists)";
    var runnerIdNote = "Identify yourself within this group?";
    if (deviceSheet) {
      var deviceHeaders = deviceSheet.getRange(1, 1, 1, deviceSheet.getLastColumn()).getValues()[0];
      // Locate the "Spreadsheet" column to pull the group pairing instructions
      var idxDevGroupCol = deviceHeaders.indexOf(COL_DEVICE.GROUP_NAME);
      if (idxDevGroupCol > -1) {
        var noteSetup = deviceSheet.getRange(1, idxDevGroupCol + 1).getNote();
        if (noteSetup && noteSetup.trim()) userSetupGuide = noteSetup.trim();
      }
      // Locate the "Runner Id" column to pull the runner profile lookup guide
      var idxDevRunnerCol = deviceHeaders.indexOf(COL_DEVICE.RUNNER_ID);
      if (idxDevRunnerCol > -1) {
        var noteRunner = deviceSheet.getRange(1, idxDevRunnerCol + 1).getNote();
        if (noteRunner && noteRunner.trim()) runnerIdNote = noteRunner.trim();
      }
    }

    // 1C. Directory Generation Loop (Populate Active Group Nodes)
    var activeDirectory = [];
    if (masterSheet) {
      // Re-use or establish the master values matrix grid cleanly
      var mData = masterSheet.getDataRange().getValues();
      var masterHeaders = mData[0];
      // Locate structural index maps via master schema definitions
      var idxReady = masterHeaders.indexOf(COL_MASTER.READY);
      var idxGroupName = masterHeaders.indexOf(COL_MASTER.GROUP_NAME);
      var idxSsId = masterHeaders.indexOf(COL_MASTER.SS_ID);
      var idxOwner = masterHeaders.indexOf(COL_MASTER.OWNER);
      // Scan directory rows to extract verified, active groups
      for (var i = 1; i < mData.length; i++) {
        var row = mData[i];
        // Ensure index boundary safety check before evaluating cell parameters
        if (idxReady > -1 && row[idxReady] !== undefined) {
          var isReadyVal = row[idxReady].toString().toUpperCase().trim();
          if (row[idxReady] === true || isReadyVal === "TRUE" || isReadyVal === "Y") {
            activeDirectory.push({
              groupName: idxGroupName > -1 ? row[idxGroupName] : "",
              ssId: idxSsId > -1 ? row[idxSsId] : "",
              ownerEmail: idxOwner > -1 ? row[idxOwner] : ""
            });
          }
        }
      }
    }

    // 1D. Cached Profile for App user, from matching their (stable) device
    var cachedProfile = null;
    if (deviceSheet) {
      var dData = deviceSheet.getDataRange().getValues();
      var deviceHeaders = dData[0];
      // Locate structural index maps via device master schema definitions
      var idxDevId = deviceHeaders.indexOf(COL_DEVICE.DEVICE_ID);
      var idxDevGroup = deviceHeaders.indexOf(COL_DEVICE.GROUP_NAME);
      var idxDevSsId = deviceHeaders.indexOf(COL_DEVICE.SS_ID);
      var idxDevRunnerId = deviceHeaders.indexOf(COL_DEVICE.RUNNER_ID);
      var idxDevResultsGid = deviceHeaders.indexOf(COL_DEVICE.RESULTS_GID);
      var idxDevNumRuns = deviceHeaders.indexOf(COL_DEVICE.COUNT);
      var idxDevAccess = deviceHeaders.indexOf(COL_DEVICE.STATUS);
      // Scan registered rows to find a matching, verified active device token
      for (var d = 1; d < dData.length; d++) {
        var dRow = dData[d];
        // Ensure index safety before checking the unique device tracking tag
          if (idxDevId > -1 && dRow[idxDevId] !== undefined
                            && dRow[idxDevId] !== "") {
          if (dRow[idxDevId].toString().trim() === devId.toString().trim()) {
            cachedProfile = {
              groupName: idxDevGroup > -1 ? dRow[idxDevGroup] : "",
              ssId: idxDevSsId > -1 ? dRow[idxDevSsId] : "",
              runnerId: idxDevRunnerId > -1 ? dRow[idxDevRunnerId] : "",
              resultsGid: idxDevResultsGid > -1 ? dRow[idxDevResultsGid] : "",
              numRuns: idxDevNumRuns > -1 ? dRow[idxDevNumRuns] : 0,
              accessStatus: idxDevAccess > -1 ? dRow[idxDevAccess] : "",
            };
            break; // Matching device confirmed
          }
        }
      }
    }
    return {
      aboutShort: aboutShort,
      aboutLong: aboutLong,
      userSetupGuideCached: userSetupGuide,
      memberRequestGuideCached: memberRequestGuide,
      runnerIdNoteCached: runnerIdNote,
      activeDirectory: activeDirectory,
      cachedProfile: cachedProfile
    };
  } catch (err) {
    return { error: err.toString() };
  }
}

/**
 * Fetch Runner Ids subset via composite mapping table
 * Filtered dynamically by the active Spreadsheet Id key
 */
function getRunnerIdsForGroup(ssIdKey) {
  try {
    if (!ssIdKey) return [];
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lookupSheet = ss.getSheetByName(TABLES.MEMBERS);
    if (!lookupSheet) return [];
    var lData = lookupSheet.getDataRange().getValues();
    var lookupHeaders = lData[0].map(function(header) {
      return header ? header.toString().trim() : "";
    });
    var idxRunnerId = lookupHeaders.indexOf(COL_MEMBER.RUNNER_ID);
    var idxSsId = lookupHeaders.indexOf(COL_MEMBER.SS_ID);
    // Safety check: If columns not found, exit cleanly
    if (idxRunnerId === -1 || idxSsId === -1) return [];
    var uniqueIds = new Set();
    // Filter rows matching the selected Group worksheet ID
    for (var i = 1; i < lData.length; i++) {
      var row = lData[i];
      var sheetSsId = row[idxSsId];
      var sheetRunnerId = row[idxRunnerId]; 
      if (sheetSsId && sheetSsId.toString().trim() === ssIdKey.toString().trim()) {
        if (sheetRunnerId && sheetRunnerId.toString().trim()) {
          uniqueIds.add(sheetRunnerId.toString().trim());
        }
      }
    }   
    // Return sorted alphabetically for clean user selection
    return Array.from(uniqueIds).sort();
  } catch (e) {
    return [];
  }
}

/**
 * Device Record Profile Submission Upsert
 */
function saveDeviceProfile(deviceId, groupName, ssId, runnerId, resultsGidMode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TABLES.DEVICES);
    if (!sheet) return { error: "Devices sheet not found" };
    
    var cleanRunnerId = runnerId ? runnerId.toString().trim() : "";
    var resultsGid = "";
    var timestamp = new Date();
    
    if (cleanRunnerId && resultsGidMode === "AUTO_LOOKUP_GID") {
      var lookupSheet = ss.getSheetByName(TABLES.MEMBERS);
      if (lookupSheet) {
        var lData = lookupSheet.getDataRange().getValues();
        var mHeaders = lData[0]; // First row contains text header values
        var idxMemberRunnerId = mHeaders.indexOf(COL_MEMBER.RUNNER_ID);
        var idxMemberSsId     = mHeaders.indexOf(COL_MEMBER.SS_ID);
        var idxMemberGid      = mHeaders.indexOf(COL_MEMBER.RESULTS_GID);
        var idxMemberNumRuns = deviceHeaders.indexOf(COL_MEMBER.COUNT);
        var idxMemberShared   = mHeaders.indexOf(COL_MEMBER.SHARED);
        for (var i = 1; i < lData.length; i++) {
          if (lData[i][idxMemberRunnerId].toString().trim() == cleanRunnerId && 
              lData[i][idxMemberSsId].toString().trim() == ssId) {
            resultsGid = lData[i][idxMemberGid];
            numRuns = lData[i][idxMemberNumRuns]; 
            var sharedRole = lData[i][idxMemberShared]
              ? lData[i][idxMemberShared].toString().trim()
              : "";
            // Execute the Access Matrix conversion rules
            if (sharedRole === "Blocked")
              accessStatus = "Needed";
            else if (sharedRole === "Editor" ||
                    sharedRole === "Viewer" ||
                    sharedRole === "Owner")
              accessStatus = "Granted";
            // else
              // accessStatus = "Requested";
            break;
          }
        }
      }
    }
    var dData = sheet.getDataRange().getValues();
    var dHeaders = dData[0]; // Devices header row array
    var idxDevTimestamp = dHeaders.indexOf(COL_DEVICE.TIMESTAMP);
    var idxDevId        = dHeaders.indexOf(COL_DEVICE.DEVICE_ID);
    var idxDevGroup     = dHeaders.indexOf(COL_DEVICE.GROUP_NAME);
    var idxDevSsId      = dHeaders.indexOf(COL_DEVICE.SS_ID);
    var idxDevRunnerId  = dHeaders.indexOf(COL_DEVICE.RUNNER_ID);
    var idxDevGid       = dHeaders.indexOf(COL_DEVICE.RESULTS_GID);
    var idxDevStatus    = dHeaders.indexOf(COL_DEVICE.STATUS);
    var matchedRow = -1;
    for (var d = 1; d < dData.length; d++) {
      if (dData[d][idxDevId] &&
          dData[d][idxDevId].toString() === deviceId.toString()) {
        matchedRow = d + 1;
        break;
      }
    }
    var deviceRowPayload = [];
    deviceRowPayload[idxDevTimestamp] = timestamp;
    deviceRowPayload[idxDevId]        = deviceId;
    deviceRowPayload[idxDevGroup]     = groupName;
    deviceRowPayload[idxDevSsId]      = ssId;
    deviceRowPayload[idxDevRunnerId]  = cleanRunnerId;
    deviceRowPayload[idxDevGid]       = resultsGid;
    deviceRowPayload[idxDevNumRuns]   = numRuns;
    deviceRowPayload[idxDevStatus]    = accessStatus;
    if (matchedRow > -1) {
      sheet.getRange(matchedRow,1,1,deviceRowPayload.length)
        .setValues([deviceRowPayload]);
    } else {
      sheet.appendRow(deviceRowPayload);
    }
    return {success: true,
            message: "Profile synchronised locally."
           };
    
  } catch (e) {
    return { error: e.toString() };
  }
}

function getTrendsRouting(ssActive,ssIdKey,runnerId) {
  var trendsList = {};
  if (runnerId) {
    var devicesSheet = ssActive.getSheetByName(TABLES.DEVICES);
    var dData = devicesSheet.getDataRange().getValues();
    var headers = dData[0]; // First row contains the text labels
    var idxSsId = headers.indexOf(COL_DEVICE.SS_ID);
    var idxRunnerId = headers.indexOf(COL_DEVICE.RUNNER_ID);
    var idxResultsGid = headers.indexOf(COL_DEVICE.RESULTS_GID);
    var idxCount = headers.indexOf(COL_DEVICE.COUNT); // skip to this row
    // Guard check to ensure all columns actually exist in the spreadsheet
    if (idxSsId !== -1 && idxRunnerId !== -1 && idxResultsGid !== -1) {
      for (var d = 1; d < dData.length; d++) {
        if (dData[d][idxSsId] == ssIdKey && dData[d][idxRunnerId] == runnerId) {
          var savedGid = dData[d][idxResultsGid];
          if (savedGid) {
            trendsList = {
              overallUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + savedGid + "&range=" + VIEWPORTS.TRENDS_OVERALL + "&viewport=focussed",
              recentUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + savedGid + "&range=" + VIEWPORTS.TRENDS_RECENT + "&viewport=focussed",
              resultsUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + savedGid + "&range=A" + idxCount.toString() + "&viewport=focussed"
            };
          }
          break;
        }
      }
    } else {
      Logger.log("Error: One or more required COL_DEVICE column headers could not be found.");
    }
  }
  return trendsList;
}

function getChartsRouting(ssActive,ssIdKey,runnerId,groupTableName) {
  var chartsDivLists = [];
  if (ssIdKey && runnerId && groupTableName) {
    if (groupTableName) {
      var groupSheet = ssActive.getSheetByName(groupTableName);
      if (groupSheet) {
        var gData = groupSheet.getDataRange().getValues();
        var gHeaders = gData[1];  // Table header displaced by seed SS ID in B1
        var idxCName = gHeaders.indexOf(COL_GROUP.CHART_NAME);
        var idxTGid = gHeaders.indexOf(COL_GROUP.TARGET_GID);
        var idxPSheet = gHeaders.indexOf(COL_GROUP.PERF_SHEET);
        var idxRange = gHeaders.indexOf(COL_GROUP.CELL_RANGE);
        var idxRIds = gHeaders.indexOf(COL_GROUP.RUNNERS_IDS);
        for (var c = 2; c < gData.length; c++) {
          var name = gData[c][idxCName];
          var targetGid = gData[c][idxTGid];
          var sectionGroup = gData[c][idxPSheet];
          var cellRange = gData[c][idxRange];
          var rawRunnerIdsList = gData[c][idxRIds] ? gData[c][idxRIds].toString() : "";
          var showLinkIcon = false;
          if (runnerId) {
            var escapedRunnerId = runnerId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');COL_MASTER
            var regex = new RegExp("(^|\\|)" + escapedRunnerId + "(\\||$)");
            showLinkIcon = regex.test(rawRunnerIdsList);
          }
          chartsDivLists.push({
            name: name,
            grouping: sectionGroup,
            showLink: showLinkIcon,
            url: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + targetGid + "&range=" + cellRange + "&viewport=focussed"
          });
        }
      }
    }
  }
  return chartsDivLists;
}

function getRankingsRouting(ssActive,ssIdKey) {
  var rankingsList = {};
  if (ssIdKey) {
    var masterSheet = ssActive.getSheetByName(TABLES.MASTER);
    var mData = masterSheet.getDataRange().getValues();
    var mHeaders = mData[0];
    var idxMasterSsId = mHeaders.indexOf(COL_MASTER.SS_ID);
    var idxMasterGroup = mHeaders.indexOf(COL_MASTER.GROUP_NAME);
    var idxLatestGid = mHeaders.indexOf(COL_MASTER.LATEST_GID);
    var idxRankingsGid = mHeaders.indexOf(COL_MASTER.RANKINGS_GID);
    var idxChallengesGid = mHeaders.indexOf(COL_MASTER.CHALLENGES_GID);
    var rankingsList = {};
    for (var i = 1; i < mData.length; i++) {
      if (mData[i][idxMasterSsId] === ssIdKey) {
        groupTableName = mData[i][idxMasterGroup];  // TODO: may need to strip & AC and perhaps Club/Parkrunner
        rankingsList = {
          latestUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + mData[i][idxLatestGid],
          currentUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + mData[i][idxRankingsGid] + "&range=" + VIEWPORTS.RANKINGS_CURRENT + "&viewport=focussed",
          bestEverUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + mData[i][idxRankingsGid] + "&range=" + VIEWPORTS.RANKINGS_BEST + "&viewport=focussed",
          challengeUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view?gid=" + mData[i][idxChallengesGid]
        };
        break;
      }
    }
  }
  return {rankingsList,groupTableName};
}

/**
 * Live Routing Dashboard Engine Mapping Routing Keys
 */
function getDashboardRouting(ssIdKey,runnerId) {
  try {
    var ssActive = SpreadsheetApp.getActiveSpreadsheet();
    var trendsList = getTrendsRouting(ssActive,ssIdKey,runnerId);
    var {rankingsList,groupTableName} = getRankingsRouting(ssActive,ssIdKey);
    var chartsDivLists = getChartsRouting(ssActive,ssIdKey,runnerId,groupTableName);
    return {
      trends: trendsList,      // Screen 1 (Left Table)  -> 3 items
      charts: chartsDivLists,  // Screen 2 (Middle Tab)  -> Filtered by division (e.g. Age Groups)
      rankings: rankingsList   // Screen 3 (Right Table) -> 4 items
    };

  } catch (e) {
    return { error: e.toString() };
  }
}

// Allows for dynamic linking of Index.html ref. to JavaScript.html
//    to include its own front-end <script> content via:
//    <?!= include('JavaScript'); ?>
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function debugRunnerLookup() {
  var testSsId = "1O7njlqIr466GiZzOGGvs9rE3t70yWMlKTyaSU7FmMzY"; // test CHAMPION Parkrunners 
  var result = getRunnerIdsForGroup(testSsId);
  Logger.log("Resulting Array: " + JSON.stringify(result));
}

function debugDashboardRouting() {
  var testSsId = "1O7njlqIr466GiZzOGGvs9rE3t70yWMlKTyaSU7FmMzY"; // test CHAMPION Parkrunners
  var testRunnerId = 'Alan_13';
  var routing = getDashboardRouting(testSsId,testRunnerId);
  Logger.log("Routing Array: " + JSON.stringify(routing));
}

function debugChartsRouting() {
  var ssActive = SpreadsheetApp.getActiveSpreadsheet();
  var testSsId = "1O7njlqIr466GiZzOGGvs9rE3t70yWMlKTyaSU7FmMzY";
  var testRunnerId = 'Alan_13';
  var testGroupName = 'CHAMPION Parkrunners';
  var routing = getChartsRouting(ssActive,testSsId,testRunnerId,testGroupName);
  Logger.log("Routing Array: " + JSON.stringify(routing));
}
