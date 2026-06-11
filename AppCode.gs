// FILE: AppCode.gs

/**
 * Global Table & Tab Configurations (Strictly Local Master Sheet Tabs)
 */
const TABLES = {
  MASTER: "CHArT5k",      // contains list of Groups with summary extracts
  DEVICES: "Devices",     // one row per App user with settings locked to their device
  RUNNERS: "Runners Gids" // composite index extracted for ALL Group spreadsheets
};

/**
 * Master CHArT5k Table Column Headers Definitions
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
  CHALLENGES_GID: "Challenges Gid"  // ordered results from planned event (same venue?)
  // Other columns include generated latest/current/ranking table links,
  //     the number of imported members, and the parkrun Group No. (if it exists)
};

/**
 * Group Sheet Table Column Headers Definitions (Referenced via Runtime Cache)
 */
const COL_GROUP = {
  CHART_NAME: "Performance Chart",  // name indicates division by Grade/Junior/Senior etc.
  SS_ID: "SS Id",                 // within unique linkable Google worksheet id,
  TARGET_GID: "Perf Gid",         // links to a single Group chart
  PERF_SHEET: "Perf Sheet",       // categorised under Age Groups, Leagues, etc. 
  CELL_RANGE: "Range",            // used to direct App user to a location on the sheet
  CHART_ID: "Chart Id",           // for highlighting App user line? (or for copying)
  RUNNERS_IDS: "Runners Ids"      // list of Runner Ids within the chart division
  // Other columns include the generated chart link,
  //     and the Chart Id in case it needs copied and tailored per user
};

/**
 * Devices Table Column Headers Definitions (required by App user for their profile)
 */
const COL_DEVICE = {
  TIMESTAMP: "Timestamp",     // A: when user first used (introduced for Web App)
  DEVICE_ID: "Device Id",     // B: device locked to App user's settings
  GROUP_NAME: "Spreadsheet",  // C: user-selected Group (with extended guide in Note)
  SS_ID: "SS Id",             // D: automatic ref from selected Group (used for linking)
  RUNNER_ID: "Runner Id",     // E: user-selected Identity (with extended guide in Note)
  RESULTS_GID: "Results Gid"  // F: automatic from Runner Id (as sheet name) in SS Id
};

/**
 * Target Cell Display Viewport Coordinate Constants
 */
const VIEWPORTS = {
  TRENDS: "N2",             // trend chart location within each runner's results sheet
  RANKINGS_CURRENT: "A8",   // current age-graded table in Rankings sheet location
  RANKINGS_BEST: "A111"     // best-ever age-graded table in Rankings sheet location
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

    // 1D. Active Device Cache Target Profile Verification Check
    // -------------------------------------------------------------
    var cachedProfile = null;
    if (deviceSheet) {
      var dData = deviceSheet.getDataRange().getValues();
      var deviceHeaders = dData[0];
      // Locate structural index maps via device master schema definitions
      var idxDevId = deviceHeaders.indexOf(COL_DEVICE.DEVICE_ID);
      var idxDevGroup = deviceHeaders.indexOf(COL_DEVICE.GROUP_NAME);
      var idxDevSsId = deviceHeaders.indexOf(COL_DEVICE.SS_ID);
      var idxDevRunner = deviceHeaders.indexOf(COL_DEVICE.RUNNER_ID);
      // Scan registered rows to find a matching, verified active device token
      for (var d = 1; d < dData.length; d++) {
        var dRow = dData[d];
        // Ensure index safety before checking the unique device tracking tag
        if (idxDevId > -1 && dRow[idxDevId] !== undefined) {
          if (dRow[idxDevId].toString().trim() === devId.toString().trim()) {
            cachedProfile = {
              groupName: idxDevGroup > -1 ? dRow[idxDevGroup] : "",
              groupSsId: idxDevSsId > -1 ? dRow[idxDevSsId] : "",
              runnerId: idxDevRunner > -1 ? dRow[idxDevRunner] : ""
            };
            break; // Match confirmed, break tracking iteration sweep early
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
    
    // 2. Check if this device is already registered in the Devices database
    var deviceSheet = ss.getSheetByName(TABLES.DEVICES);
    var cachedProfile = null;
    
    if (deviceSheet) {
      var devData = deviceSheet.getDataRange().getValues();
      for (var i = 1; i < devData.length; i++) {
        if (devData[i][1] && devData[i][1].toString() === devId.toString()) { // Column B: Device ID
          cachedProfile = {
            groupName: devData[i][2], // Column C: Group Name
            groupSsId: devData[i][3], // Column D: Linked Spreadsheet ID
            runnerId: devData[i][4]   // Column E: Runner Token
          };
          break;
        }
      }
    }
    
    return {
      aboutShort: aboutShort,
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
 * Fetch Runner IDs subset from central local composite mapping table
 */
function getRunnerIdsForGroup(ssIdKey) {
  try {
    if (!ssIdKey) return [];
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lookupSheet = ss.getSheetByName(TABLES.RUNNERS);
    if (!lookupSheet) return [];

    var lData = lookupSheet.getDataRange().getValues();
    var uniqueIds = new Set();
    
    for (var i = 1; i < lData.length; i++) {
      var sheetRunnerId = lData[i][0];
      var sheetSsId = lData[i][1];
      
      if (sheetSsId && sheetSsId.toString() === ssIdKey.toString()) {
        if (sheetRunnerId && sheetRunnerId.toString().trim()) {
          uniqueIds.add(sheetRunnerId.toString().trim());
        }
      }
    }
    
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
      var lookupSheet = ss.getSheetByName(TABLES.RUNNERS);
      if (lookupSheet) {
        var lData = lookupSheet.getDataRange().getValues();
        for (var i = 1; i < lData.length; i++) {
          if (lData[i][0].toString().trim() == cleanRunnerId && lData[i][1].toString().trim() == ssId) {
            resultsGid = lData[i][2]; 
            break;
          }
        }
      }
    }

    var data = sheet.getDataRange().getValues();
    var matchedRow = -1;
    for (var d = 1; d < data.length; d++) {
      if (data[d][1] && data[d][1].toString() === deviceId.toString()) {
        matchedRow = d + 1;
        break;
      }
    }

    if (matchedRow > -1) {
      sheet.getRange(matchedRow, 1, 1, 6).setValues([[timestamp, deviceId, groupName, ssId, cleanRunnerId, resultsGid]]);
    } else {
      sheet.appendRow([timestamp, deviceId, groupName, ssId, cleanRunnerId, resultsGid]);
    }
    
    return { success: true, message: "Profile synchronized locally." };
    
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Live Routing Dashboard Engine Mapping Routing Keys
 */
function getDashboardRouting(ssIdKey, runnerId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(TABLES.MASTER);
    var mData = masterSheet.getDataRange().getValues();
    var mHeaders = mData[0];
    
    var idxMasterSsId = mHeaders.indexOf(COL_MASTER.SS_ID);
    var idxMasterGroup = mHeaders.indexOf(COL_MASTER.GROUP_NAME);
    var idxLatestGid = mHeaders.indexOf(COL_MASTER.LATEST_GID);
    var idxRankingsGid = mHeaders.indexOf(COL_MASTER.RANKINGS_GID);

    var rankingsData = {};
    var localTabName = "";
    
    for (var i = 1; i < mData.length; i++) {
      if (mData[i][idxMasterSsId] === ssIdKey) {
        localTabName = mData[i][idxMasterGroup];
        rankingsData = {
          latestUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view#gid=" + mData[i][idxLatestGid],
          currentUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view#gid=" + mData[i][idxRankingsGid] + "&range=" + VIEWPORTS.RANKINGS_CURRENT,
          bestEverUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view#gid=" + mData[i][idxRankingsGid] + "&range=" + VIEWPORTS.RANKINGS_BEST
        };
        break;
      }
    }

    var trendsUrl = "";
    if (runnerId) {
      var devicesSheet = ss.getSheetByName(TABLES.DEVICES);
      var dData = devicesSheet.getDataRange().getValues();
      for (var d = 1; d < dData.length; d++) {
        if (dData[d][3] == ssIdKey && dData[d][4] == runnerId) {
          var savedGid = dData[d][5];
          if (savedGid) {
            trendsUrl = "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view#gid=" + savedGid + "&range=" + VIEWPORTS.TRENDS;
          }
          break;
        }
      }
    }

    var chartsList = [];
    if (localTabName) {
      var groupSheet = ss.getSheetByName(localTabName);
      if (groupSheet) {
        var gData = groupSheet.getDataRange().getValues();
        var gHeaders = gData[0];
        
        var idxCName = gHeaders.indexOf(COL_GROUP.CHART_NAME);
        var idxTGid = gHeaders.indexOf(COL_GROUP.TARGET_GID);
        var idxPSheet = gHeaders.indexOf(COL_GROUP.PERF_SHEET);
        var idxRange = gHeaders.indexOf(COL_GROUP.CELL_RANGE);
        var idxRIds = gHeaders.indexOf(COL_GROUP.RUNNER_IDS);
        
        for (var c = 1; c < gData.length; c++) {
          var name = gData[c][idxCName];
          var targetGid = gData[c][idxTGid];
          var sectionGroup = gData[c][idxPSheet];
          var cellRange = gData[c][idxRange];
          var rawRunnerIdsList = gData[c][idxRIds] ? gData[c][idxRIds].toString() : "";
          
          var showLinkIcon = false;
          if (runnerId) {
            var escapedRunnerId = runnerId.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
            var regex = new RegExp("(^|\\|)" + escapedRunnerId + "(\\||$)");
            showLinkIcon = regex.test(rawRunnerIdsList);
          }

          chartsList.push({
            name: name,
            grouping: sectionGroup,
            showLink: showLinkIcon,
            url: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view#gid=" + targetGid + "&range=" + cellRange
          });
        }
      }
    }

    return {
      trendsUrl: trendsUrl,
      rankings: rankingsData,
      charts: chartsList
    };

  } catch (e) {
    return { error: e.toString() };
  }
}

function debugRunnerLookup() {
  var testSsId = "YOUR_REAL_TEST_SSID_STRING"; 
  var result = getRunnerIdsForGroup(testSsId);
  Logger.log("Resulting Array: " + JSON.stringify(result));
}
