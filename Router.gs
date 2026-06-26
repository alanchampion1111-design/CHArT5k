// 1. Collated Distribution Router (Connects to CHArT5k Planet Hub)

// 1a. All menu entry points (Ctrl-0..Ctrl-9)
function ImportResultForEachRunner() {    // also a trigger
  CHArT5kPlanet.ImportResultForEachRunner();
}
function ImportResultsOnEventDate() {
  CHArT5kPlanet.ImportResultsOnEventDate();
}
function AppendNonParkrunResult() {
  CHArT5kPlanet.AppendNonParkrunResult();
}
function CatchUpAllPositions() {    // also a trigger
  CHArT5kPlanet.CatchUpAllPositions();
}
function ReprotectResultsSheetPerRunner() {
  CHArT5kPlanet.ReprotectResultsSheetPerRunner();
}
function ColourLegendsInGroups() {
  CHArT5kPlanet.ColourLegendsInGroups();
}
function GenBatchChartsFromGroups() {
  CHArT5kPlanet.GenBatchChartsFromGroups();
}
function AddNewMember() {
  CHArT5kPlanet.AddNewMember();
}
function DeleteExistingMember() {
  CHArT5kPlanet.DeleteExistingMember();
}
function SpawnNewGroup() {
  CHArT5kPlanet.SpawnNewGroup();
}

// 1b. Trigger entry points
function BatchPositionsForRunner() {    // needs properties
  CHArT5kPlanet.BatchPositionsForRunner();
}
function GenerateAgeGroupCharts() {     // uses index on property
  CHArT5kPlanet.GenerateAgeGroupCharts();
}
function GenerateLeagueGroupCharts() {
  CHArT5kPlanet.GenerateLeagueGroupCharts();
}
function GenerateGenderGroupCharts() {
  CHArT5kPlanet.GenerateGenderGroupCharts();
}
function GenerateFamiliesGroupCharts() {
  CHArT5kPlanet.GenerateFamiliesGroupCharts();
}
function PrepareAppSheets() {
  CHArT5kPlanet.PrepareAppSheets();
}

// 1c. Extra Macro entry points
function AcceptCookies() {
  CHArT5kPlanet.AcceptCookies();
}

// 2. Universal UI/Trigger/Macro Router
// Routes all custom menu clicks straight into the Library's function tree.
function ActionRouter(e) {
  var functionToRun = e ? e.controlName : null;
  if (functionToRun && typeof CHArT5kPlanet[functionToRun] === 'function') {
    CHArT5kPlanet[functionToRun]();
  }
}

// 3. onOpen functions
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
      'AddNewMember')
    .addItem("Delete existing member"+
      "\u00A0".repeat(24)+"Ctrl+Alt+Shift+8",
      'DeleteExistingMember')
    .addItem("Spawn new club or family"+
      "\u00A0".repeat(21)+"Ctrl+Alt+Shift+9",
      'SpawnNewGroup')
    // .insertMenu(ui,5)   // ideally before Tools 
    .addToUi();
}
