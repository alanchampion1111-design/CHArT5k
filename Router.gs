// 1. Collated Distribution Router (Connects to sources mastered on CHArT5k Planet Hub)

// 1a. All menu entry points (Ctrl-0..Ctrl-9)
function ImportResultForEachRunner(eventDate) {    // also a scheduled trigger (default eventDate)
  CHArT5kPlanet.ImportResultForEachRunner(eventDate);
}
function ImportResultsOnEventDate() {              // prompts for date from user to call 1a. ImportResult...
  CHArT5kPlanet.ImportResultsOnEventDate();
}
function AppendNonParkrunResult() {              // creates a new result row for manual entry
  CHArT5kPlanet.AppendNonParkrunResult();
}
function CatchUpAllPositions() {                  // also triggered from 1d. DoAdd... & DoSpawn...
  CHArT5kPlanet.CatchUpAllPositions();
}
function ReprotectResultsSheetPerRunner() {        // also called from 1d. OnEdit...
  CHArT5kPlanet.ReprotectResultsSheetPerRunner();
}
function ColourLegendsInGroups() {                // potentially also called from GenBatch...
  CHArT5kPlanet.ColourLegendsInGroups();
}
function GenBatchChartsFromGroups() {            // also a scheduled trigger after (series of) ImportResult...
  CHArT5kPlanet.GenBatchChartsFromGroups();
}
function AddNewMember() {                // form entry to callback 1d. DoAddNewMember(formData)
  CHArT5kPlanet.AddNewMember();
}
function DeleteExistingMember() {        // form entry to  callback1d. DoDeleteExistingMember(formData)
  CHArT5kPlanet.DeleteExistingMember();
}
function SpawnNewGroup() {              // form entry to callback 1d. DoSpawnNewGroup(formData)
  CHArT5kPlanet.SpawnNewGroup();
}

// 1b. Triggered entry points
function ReImportResultForEachRunner(eventDate) {    // called from 1a. ImportResult... (and needs cleaned)
  CHArT5kPlanet.ReImportResultForEachRunner(eventDate);
}
function BatchPositionsForRunner() {         // called from 1a. CatchUp... needs properties (and needs cleaned)
  CHArT5kPlanet.BatchPositionsForRunner();
}
function GenerateAgeGroupsBlockCharts() {    // called from GenBatch... as a series with index on sheet (and needs cleaned)
  CHArT5kPlanet.GenerateAgeGroupsBlockCharts();
}
function GenerateLeaguesBlockCharts() {      // called from GenBatch... as a series with index on sheet (and needs cleaned)
  CHArT5kPlanet.GenerateLeaguesBlockCharts();
}
function GenerateGenderBlockCharts() {       // called from GenBatch... as a series with index on sheet (and needs cleaned)
  CHArT5kPlanet.GenerateGenderBlockCharts();
}
function GenerateFamiliesBlockCharts() {     // called from GenBatch... as a series with index on sheet (and needs cleaned)
  CHArT5kPlanet.GenerateFamiliesBlockCharts();
}
function PrepareAppSheets() {                // Scheduled trigger to export extract (after import & generate) for App to pull
  CHArT5kPlanet.PrepareAppSheets();
}

// 1c. Extra Macro entry points
function AcceptCookies() {                // If required (under stealth), perhaps once per six moths?
  CHArT5kPlanet.AcceptCookies();
}

// 1d. UI Callback entry points
function onOpen(event) {                      // built-in Google default here and in the Hub 
  CHArT5kPlanet.onOpen(event);                // Assumes that contains:  var ui = SpreadsheetApp.getUi();
}
function onEditDetectNeedToReprotect(event) {        // callback from 1a. AppendNonParkrunResult
  CHArT5kPlanet.onEditDetectNeedToReprotect(event);
}
function DoSpawnNewGroup(formData) {                 // callback from 1a. SpawnNewGroup
  CHArT5kPlanet.DoSpawnNewGroup(formData);
}
function DoAddNewMember(formData) {                  // callback from 1a. AddNewMember
  CHArT5kPlanet.DoAddNewMember(formData);
}
function DoDeleteExistingMember(formData) {          // callback from 1a. DeleteExistingMember
  CHArT5kPlanet.DoDeleteExistingMember(formData);
}

// 2. Universal UI/Trigger/Macro Router
// Routes all custom menu clicks straight into the Library's function tree.
function ActionRouter(e) {
  var functionToRun = e ? e.controlName : null;
  if (functionToRun && typeof CHArT5kPlanet[functionToRun] === 'function') {
    CHArT5kPlanet[functionToRun]();
  }
}
