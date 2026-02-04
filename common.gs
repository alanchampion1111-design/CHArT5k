// Common global settings available to local funtions and GCR functions in separate files

// used to track the main spreadsheet operation which may flip
// from current to a new context after creating a new family instance
const debug = false;   // WARNING: debug if true may slow down performance and may skip runners!!
var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();  // allows dynamic shift to new context
var activeSpreadsheetId = activeSpreadsheet.getId();   // the originator ID
if (debug) Logger.log('Current Spreadsheet Id: '+activeSpreadsheetId);
const runnersSheetNAME = 'Runners';
var allRunnersSheet = activeSpreadsheet   // MUST redo after a dynamic shift during spawning
  .getSheetByName(runnersSheetNAME);     // ...for new runners in existing 
const templateNameCELL = "J1";          // This cell identifes the seed template (e.g. Keren)
const templateNAME = allRunnersSheet     // The seed template may be readily reconfigured
  .getRange(templateNameCELL)     // See the note on Runners!J1 cell
  .getValue();                    //  ...where only the first 3 (or 4) rows are relevant
const templateSPREADSHEET = 'FAMILY Template';
const templateFOLDER ='Spawned';
const clubTYPE = 'Parkrunners';   // alternatively, 'ClubRunners'?

const browserURL = 'https://browser-automation-service-224251628103.europe-west1.run.app';    // Google Cloud service in operation
// const sampleURL = 'https://www.example.com';  // default test
const sampleURL = 'https://www.parkrun.org.uk/colchestercastle/results/116'
var browserSession;   // Shared by threaded/recursed processes for up to 30 minutes because browser lingers

// used for import,etc,
const parkrunnerIdCOL = 10;       // in column J on Runners sheet
const parkrunnerIdINDEX = 9;      // in column J on Runners sheet (for arrays or range offsets)
const runnersStartROW = 3;        // start after title & header rows (2)
const resultsStartROW = 3;        // start after title & header rows (2)
const numBlankROWS = 1;           // Regular catch-up of result by those permitted
const dateFORMAT = 'd-MMM-yy';    //consistent for backwards compatibility
const batchSizeMAX = 10;    // estimate batch to catch up on within 5-6 minutes
