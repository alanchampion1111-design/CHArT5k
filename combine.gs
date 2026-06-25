// Common global settings available to local funtions and GCR functions in separate files

// used to track the main spreadsheet operation which may flip
// from current to a new context after creating a new family instance
const debug = false;   // WARNING: debug if true may slow down performance and may skip runners!!
var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();  // allows dynamic shift to new context
var activeSpreadsheetId = activeSpreadsheet.getId();   // the originator ID
if (debug) Logger.log('Current Spreadsheet Id: '+activeSpreadsheetId);
const runnersSheetNAME = 'Runners';
var allRunnersSheet = activeSpreadsheet   // MUST redo after a dynamic shift during spawning
  .getSheetByName(runnersSheetNAME);      // ...for new runners from initiating Spreadsheet 
const importDateCELL = "I1";              // Runners cell for the last import date (d-MMM-yy)
const clubIdCELL = "J1";                  // Runners cell identifies the club or family group id
const importIndexCELL = "K1";             // Runners cell with index of runner to continue import
const importTotalCELL = "L1";             // Runners cell with number of runners on import date
const templateNameCELL = "E1";            // Runners cell identifies the seed template (e.g. Joe_90)
const templateNAME = allRunnersSheet      // The seed template may be readily reconfigured
  .getRange(templateNameCELL)     // See the note on Runners!J1 cell
  .getValue();                    //  ...where only the first 3 (or 4) rows are relevant

const browserURL = 'https://browser-automation-service-224251628103.europe-west1.run.app';    // Google Cloud service in operation
// const sampleURL = 'https://www.example.com';  // default test
const sampleURL = 'https://www.parkrun.org.uk/colchestercastle/results/116'
var browserSession;   // Shared by threaded/recursed processes for up to 30 minutes because browser lingers

// used for import,etc,
const runnersStartROW = 3;        // start after title & header rows (2)
const dobINDEX = 4;               // in column E on Runners sheet (for arrays or range offsets)
const parkrunnerIdINDEX = 9;      // in column J on Runners sheet (for arrays or range offsets)
const parkrunnerIdCOL = 10;       // in column J on Runners sheet
const resultsStartROW = 3;        // start after title & header rows (2)
const PBtickBoxCOL = 8;           // column H tick-box on Results sheets, ticked if value in G is PB
const genderPosnCOL = 9;          // column I on each Results sheet (until on parkrun results page?)
const ageCatPosnCOL = 10;         // column J on each Results sheet
const ageGradePosnCOL = 11;       // column K on each Results sheet
const ageCatCOL = 12;             // column L is the Age Category on event date (derived from DoB)
const PBtoAgeCatNumCOLS = 5;      // cols H..L for I..K positions between derived cols, H & L
const maxChangeROWS = 3;          // restrict changes to recent results
const groupPerfsCOL = 4;          // column D on each Performances chart sheet (e.g. Leagues)
const chartColOFFSET = -2;        // column B from D on any Performances chart sheet
const trendColOFFSET = 12;        // column N from B on runner's own trend chart sheet
const dateFORMAT = 'd-MMM-yy';    // consistent for backwards compatibility
const batchSizeMAX = 13;          // 1/4 year - estimate batch to catch up on within 5-6 minutes

const runnerNameCOLUMN = "A";     // Runners name in column A 
const runnerSurnameCOLUMN = "B";  // Runners surname in column B 
const parkrunnerIdCOLUMN = "J";   // Runners parkrunner ID in column J
const hasResultsCOLUMN = "K";     // ...results exist (D3:D), with Parkrunner Id in col J
const hasPosnsCOLUMN = "L";       // ...has Positions up-to-date (I3:I) based on genderPosnCOL

/**
 * Sets up the Parkruns menu on opening the spreadsheet.
 *  Instructions to set up (and execute?) the trigger:
 *    1. Go to the Apps Script (Macros) editor.
 *    2. Click on the clock icon (Triggers) in the left sidebar.
 *    3. Click on "Create trigger".
 *    4. Set up the trigger with the following settings:
 *        - Choose function: onOpen
 *        - Select event type: On open
 *        - Save
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  var parkrunsMenu = ui.createMenu('parkrun')
    .addItem("Import result for each runner"+
      "\u00A0".repeat(16)+"Ctrl+Alt+Shift+0",
      'ImportResultForEachRunner')
    .addItem("Import results on event date"+
      "\u00A0".repeat(17)+"Ctrl+Alt+Shift+1",
      'ImportResultsOnEventDate')
    .addItem("Append non-parkrun result"+
      "\u00A0".repeat(20)+"Ctrl+Alt+Shift+2",
      "AppendNonParkrunResult")
    .addItem("Catch-up all positions"+
      "\u00A0".repeat(28)+"Ctrl+Alt+Shift+3",
      'CatchUpAllPositions')
    .addSeparator()
    .addItem("Protect results sheet per runner"+
      "\u00A0".repeat(11)+"Ctrl+Alt+Shift+4",
      'ReprotectResultsSheetPerRunner')
    .addItem("Colour legends in Groups"+
      "\u00A0".repeat(22)+"Ctrl+Alt+Shift+5",
      'ColourLegendsInGroups')
    .addItem("Generate charts from Groups"+
      "\u00A0".repeat(15)+"Ctrl+Alt+Shift+6",
      'GenBatchChartsFromGroups')
    .addSeparator()
    .addItem("Add new member"+
      "\u00A0".repeat(35)+"Ctrl+Alt+Shift+7",
      'AddFamilyMember')
    .addItem("Delete existing member"+
      "\u00A0".repeat(24)+"Ctrl+Alt+Shift+8",
      'DeleteFamilyMember')
    .addItem("Spawn new club or family"+
      "\u00A0".repeat(21)+"Ctrl+Alt+Shift+9",
      'SpawnNewGroup')
    // .insertMenu(ui,5)   // ideally before Tools 
    .addToUi();
}
