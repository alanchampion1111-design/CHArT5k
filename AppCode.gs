/**
 * Global Table & Tab Configurations (Strictly Local Master Sheet Tabs)
 */
const TABLES = {
  MASTER: "CHArT5k",
  DEVICES: "Devices",
  RUNNERS: "Runners Gids"
};

/**
 * Master CHArT5k Table Column Headers Definitions
 */
const COL_MASTER = {
  READY: "Ready",
  GROUP_NAME: "Spreadsheet",
  SS_ID: "SS Id",
  OWNER: "Owner",
  LATEST_GID: "Latest Gid",
  RANKINGS_GID: "Rankings Gid"
};

/**
 * Group Sheet Table Column Headers Definitions (Referenced via Runtime Cache)
 */
const COL_GROUP = {
  CHART_NAME: "Performance Chart",
  TARGET_GID: "Perf Gid",
  PERF_SHEET: "Perf Sheet",
  CELL_RANGE: "Range",
  RUNNER_IDS: "Runners Ids"
};

/**
 * Target Cell Display Viewport Coordinate Constants
 */
const VIEWPORTS = {
  TRENDS: "N2",
  RANKINGS_CURRENT: "A8",
  RANKINGS_BEST: "A111"
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
 * STAGE 1A: App Initialization
 * Parses local Master directory, reads header Notes, and queries current device profile state.
 * Strictly Read-Only for infrastructure sheets; no external files are scanned.
 */
function getAppInitializationData(localDeviceId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var masterSheet = ss.getSheetByName(TABLES.MASTER);
    var devicesSheet = ss.getSheetByName(TABLES.DEVICES);
    
    var masterData = masterSheet.getDataRange().getValues();
    var headers = masterData[0];
    
    var idxReady = headers.indexOf(COL_MASTER.READY);
    var idxGroup = headers.indexOf(COL_MASTER.GROUP_NAME);
    var idxSsId = headers.indexOf(COL_MASTER.SS_ID);
    var idxOwner = headers.indexOf(COL_MASTER.OWNER);

    var groupsList = [];
    
    // Default Fallback Guidelines Text UI
    var aboutShort = "CHArT5k System Ready";
    var aboutFull = "Operational tracking framework.";
    var userSetupGuide = "Please select your existing club or family Group and your existing Runner Id within that Group.";
    var memberRequestGuide = "Please select your existing club or family Group.";
    var runnerIdNote = "Please select your unique tracking ID assigned within this specific group.";

    for (var i = 1; i < masterData.length; i++) {
      var row = masterData[i];
      if (row[idxReady] === true || row[idxReady] === "TRUE") {
        groupsList.push({
          groupName: row[idxGroup],
          ssId: row[idxSsId],
          ownerEmail: row[idxOwner]
        });
      }
    }

    if (masterData.length > 1) {
      if (masterSheet.getRange(1, idxReady + 1).getNote()) aboutShort = masterSheet.getRange(1, idxReady + 1).getNote();
      if (masterSheet.getRange(1, idxGroup + 1).getNote()) aboutFull = masterSheet.getRange(1, idxGroup + 1).getNote();
    }

    // Capture dynamic setup instructions directly from the Devices tab column header notes
    if (devicesSheet) {
      var devHeaders = devicesSheet.getDataRange().getValues()[0];
      var idxDevGroup = devHeaders.indexOf("Spreadsheet");
      var idxDevRunner = devHeaders.indexOf("Runner Id");
      
      if (idxDevGroup > -1 && devicesSheet.getRange(1, idxDevGroup + 1).getNote()) {
        userSetupGuide = devicesSheet.getRange(1, idxDevGroup + 1).getNote();
      }
      if (idxDevRunner > -1 && devicesSheet.getRange(1, idxDevRunner + 1).getNote()) {
        runnerIdNote = devicesSheet.getRange(1, idxDevRunner + 1).getNote();
      }
    }

    var deviceProfile = null;
    if (devicesSheet && localDeviceId) {
      var devData = devicesSheet.getDataRange().getValues();
      for (var d = 1; d < devData.length; d++) {
        // Match checking column 1 (Index index 1) for Saved Device ID Tag
        if (devData[d][1] && devData[d][1].toString() === localDeviceId.toString()) {
          deviceProfile = {
            groupName: devData[d][2],
            ssId: devData[d][3],
            runnerId: devData[d][4]
          };
          break;
        }
      }
    }

    return {
      aboutShort: aboutShort,
      aboutFull: aboutFull,
      userSetupGuide: userSetupGuide,
      memberRequestGuide: memberRequestGuide,
      runnerIdNote: runnerIdNote,
      activeGroups: groupsList,
      deviceProfile: deviceProfile
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * STAGE 1B: Fetch Runner IDs subset from central local composite mapping table
 * Evaluates entirely against local internal "Runners Gids" tab row sets.
 */
function getRunnerIdsForGroup(ssIdKey) {
  try {
    if (!ssIdKey) return [];
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lookupSheet = ss.getSheetByName(TABLES.RUNNERS);
    if (!lookupSheet) return [];

    var lData = lookupSheet.getDataRange().getValues();
    var uniqueIds = new Set();
    
    // Scans local mapping sheet: Column A (Runner ID), Column B (Group SS ID)
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
 * STAGE 1C: Device Record Profile Submission Upsert
 * Write access is restricted solely to the "Devices" tab.
 */
function saveDeviceProfile(deviceId, groupName, ssId, runnerId, resultsGidMode) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(TABLES.DEVICES);
    if (!sheet) return { error: "Devices sheet not found" };
    
    var cleanRunnerId = runnerId ? runnerId.toString().trim() : "";
    var resultsGid = "";
    var timestamp = new Date();
    
    // Search the local composite mapping table ONLY if handling an Existing Member with an ID
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

    // Database payload structure configuration mapping columns:
    // [Timestamp, Device ID, Spreadsheet (Group), SS Id, Runner Id, Results Gid]
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
 * STAGE 1D: Live Routing Dashboard Engine Mapping Routing Keys
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
          currentUrl: "https://docs.google.com/spreadsheets/d/" + ssIdKey + "/view#gid=" + ssIdKey + "/view#gid=" + mData[i][idxRankingsGid] + "&range=" + VIEWPORTS.RANKINGS_CURRENT,
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
