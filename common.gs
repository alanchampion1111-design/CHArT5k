// Common global settings available to local funtions and GCR functions in separate files

// used to track the main spreadsheet operation which may flip
// from current to a new context after creating a new family instance
let activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();  // allows dynamic shift to new context
const runnersSheetNAME = 'Runners';
let allRunnersSheet = activeSpreadsheet   // MUST redo after a dynamic shift during spawning
  .getSheetByName(runnersSheetNAME);     // ...for new runners in existing 
const templateNameCELL = "J1";          // This cell identifes the seed template (e.g. Keren)
const templateNAME = allRunnersSheet     // The seed template may be readily reconfigured
  .getRange(templateNameCELL)     // See the note on Runners!J1 cell
  .getValue();                    //  ...where only the first 3 (or 4) rows are relevant
const curSpreadsheetID = activeSpreadsheet.getId();   // the originator ID
let famSpreadsheetId;             // used if a new unique instance of the spreadsheet
const templateSPREADSHEET = 'FAMILY Template';
const templateFOLDER ='Spawned';
const clubTYPE = 'Parkrunners';   // or 'ClubRunners'

// used for import,etc,
const parkrunnerIdCOL = 10;       // in column J on Runners sheet
const parkrunnerIdINDEX = 9;      // in column J on Runners sheet (for arrays or range offsets)
const runnersStartROW = 3;        // start after title & header rows (2)
const resultsStartROW = 3;        // start after title & header rows (2)
const numBlankROWS = 1;           // Regular catch-up of result by those permitted
const dateFORMAT = 'd-MMM-yy';    //consistent for backwards compatibility
