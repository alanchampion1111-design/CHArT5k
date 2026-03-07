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
const importDateCELL = "I1";              // Runners cell for the last import date (dd/MM/yyyy)
const importIndexCELL = "K1";             // Runners cell with index of runner to continue import
const importTotalCELL = "L1";             // Runners cell with number of runners on import date
const templateNameCELL = "J1";            // Runners cell identifies the seed template (e.g. Joe_90)
const templateNAME = allRunnersSheet      // The seed template may be readily reconfigured
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
const runnersStartROW = 3;        // start after title & header rows (2)
const parkrunnerIdINDEX = 9;      // in column J on Runners sheet (for arrays or range offsets)
const parkrunnerIdCOL = 10;       // in column J on Runners sheet
const resultsStartROW = 3;        // start after title & header rows (2)
const genderPosnCOL = 9;          // column I on each runner Results sheet (until on parkrun results page?)
const ageCatPosnCOL = 10;         // column J on each runner Results sheet
const ageGradePosnCOL = 11;       // column K on each runner Results sheet
const ageCatCOL = 12;             // column L is the Age Category on event date (derived from DoB)
const PBtoAgeCatNumCOLS = 5;      // cols H..L for I..K positions between derived cols, H & L
const numBlankROWS = 1;           // Regular catch-up of result by those permitted
const dateFORMAT = 'd-MMM-yy';    // consistent for backwards compatibility
const batchSizeMAX = 13;          // 1/4 year - estimate batch to catch up on within 5-6 minutes
