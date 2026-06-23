// FILE: PullCode.gs

DEBUG = true;
/**
 * Global configurations for the Weekly Data Pull Engine.
 * Running this via a weekly time-driven trigger refreshes the CHArT5k "Moon" workspace.
 */
const masterCONFIG = {
  directorySHEET: "CHArT5k",
  allMembersSHEET: "Members",
  // Chart sheets exist per Group
  COL_SS_ID: "SS Id",
  COL_GROUP: "Spreadsheet",
  COL_READY: "Ready"
};

/**
 * Executed via (weekly) trigger after ALL "planet" Groups tasks are completed:
 *  A. import of each Group member's results
 *  B. generation of Group divisional charts
 *  C. preparation for App by pushing to Overview, Trends and Charts sheets
 * Source -> Target sheets for pulling extracts into this "moon" App worksheet are:
 *  1. Overview -> CHArT5k directory (one row per Group) incl. Rankings sheet refs.
 *  2. Trends -> Members extract (one row per runner across ALL Groups)
 *  3. Charts -> <Group name> (e.g. CHAMPION Parkrunners (one row per comparative chart)
 */
function pullExtractFromGroups() {
  var ssMoon = SpreadsheetApp.getActiveSpreadsheet();
  var directorySheet = ssMoon.getSheetByName(masterCONFIG.directorySHEET);
  if (!directorySheet) {  // overkill since sheet name unchanged?
    throw new Error("Group SS directory sheet '" + masterCONFIG.directorySHEET + "' not found in CHArT5k worksheet.");
  }
  // Extract all rows from the directory to locate our planet targets
  var dirData = directorySheet.getDataRange().getValues();
  var dirHeaders = dirData[0];
  var idxSsId = dirHeaders.indexOf(masterCONFIG.COL_SS_ID);
  var idxGroup = dirHeaders.indexOf(masterCONFIG.COL_GROUP);
  var idxReady = dirHeaders.indexOf(masterCONFIG.COL_READY);
  if (idxSsId === -1)
    throw new Error("Could not find 'SS Id' column in the directory headers.");
  var directoryRowsUpdate = [];
  var compositeTrendsRows = [];
  if (DEBUG) Logger.log("Extract for each Group...");

  // For each Group SS row in directory if Ready?
  for (var i = 2; i < dirData.length; i++) {  // skip header and template rows
    var dirRow = dirData[i];
    var isReady = idxReady > -1
      ? String(dirRow[idxReady]).toUpperCase()
      : "undefined";
    if (idxReady > -1 && isReady !== "TRUE")
      continue; // skip groups not ticked as Ready
    var planetId = dirRow[idxSsId];
    if (!planetId) continue;
    var planetGroup = dirRow[idxGroup];
    if (!planetGroup) continue;
    try { // extract from each Group SS "planet"
      var ssPlanet = SpreadsheetApp.openById(planetId);
       if (DEBUG) Logger.log("Group: "+planetGroup);

      // 1: Extract & Transpose Overview -> CHArT5k (as directory)
      var overviewSheet = ssPlanet.getSheetByName("Overview");
      if (overviewSheet) {
        var ovData = overviewSheet.getDataRange().getValues();
        if (DEBUG)
          Logger.log("1A. Transposing "+(ovData.length-1)+
            " values from 'Overview' for Group: "+planetGroup);
        var valueVector = [isReady]; // include state in 1st column
        for (var k=1; k<ovData.length; k++) { // skip header
          valueVector.push(ovData[k][1]); 
        }
        var targetRowNumber = i+1;  // ensure updating same row!
        directorySheet
          .getRange(targetRowNumber,1,1,valueVector.length)
          .setValues([valueVector]);
        if (DEBUG)
          Logger.log("1B. In-situ update completed for row, " +   targetRowNumber + " (" + planetGroup + ")");
          
      }
      
      // 2A: Extract Trends -> Composite Members ('Members')
      var trendsSheet = ssPlanet.getSheetByName("Trends");
      if (trendsSheet) {
        var trendData = trendsSheet.getDataRange().getValues();
        if (trendData.length > 1) {
          var rawTrends = trendData.slice(2); // ignore header and extra row?
          compositeTrendsRows = compositeTrendsRows.concat(rawTrends);
        }
        if (DEBUG)
          Logger.log("1C. Extract from 'Trends' for Group: "+planetGroup);
      }
      
      // 3: Extract Charts -> Clean Group-Named Sheet
      var chartsSheet = ssPlanet.getSheetByName("Charts");
      if (chartsSheet) {
        var chartData = chartsSheet.getDataRange().getValues();
        if (DEBUG)
          Logger.log("3A. Extract from 'Charts' for Group: "+planetGroup);
        var rawPlanetName = ssPlanet.getName(); 
        var cleanSheetName = rawPlanetName.replace(/\s*&\s*AC\b/i, "").trim(); 
        var moonChartTarget = ssMoon.getSheetByName(cleanSheetName);
        if (!moonChartTarget) { // update Group sheet; do not replace
          moonChartTarget = ssMoon.insertSheet(cleanSheetName);
        } // missing header if new!
        if (moonChartTarget.getLastRow() > 2) {
          moonChartTarget.getRange(3,1,moonChartTarget.getLastRow()-2, moonChartTarget.getLastColumn()).clearContent();
        }
        // 🔹 Write full dataset safely using range bounds
        moonChartTarget.getRange(1,1,chartData.length,chartData[0].length)
          .setValues(chartData);
        if (DEBUG)
          Logger.log("3B. Sync 'Charts' for Group: "+rawPlanetName+
          " to sheet, "+cleanSheetName);
      }
      
    } catch (err) {
      Logger.log("Error extracting from Group SS (" + planetGroup + "): " + err.toString());
    }
  }
  
  // 2B. Commit update to target composite for all Members
  if (compositeTrendsRows.length > 0) {
    var destComposite = ssMoon.getSheetByName(masterCONFIG.allMembersSHEET);
    if (destComposite) {
      if (destComposite.getLastRow() > 1) {
        destComposite.getRange(2, 1, destComposite.getLastRow() - 1, destComposite.getLastColumn()).clearContent();
      }
      destComposite.getRange(2, 1, compositeTrendsRows.length, compositeTrendsRows[0].length).setValues(compositeTrendsRows);
      if (DEBUG)
        Logger.log("2B. Sync members for ALL Groups to sheet: " + masterCONFIG.allMembersSHEET);
    } else {
      Logger.log("ERROR: Could not find sheet named '" + masterCONFIG.allMembersSHEET + "' inside the Moon worksheet!");
    }
  }
  Logger.log("Data extraction and synchronization executed successfully.");
}
