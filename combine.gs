// Combined menu settings for:
//    local functions (Ctrl+Alt+Shift+2) and (Ctrl+Alt+Shift+4..6)
//    GCR functions (Ctrl+Alt+Shift+0..1), (Ctrl+Alt+Shift+0..3) and (Ctrl+Alt+Shift+7..9)
// ...although DeleteExistingMember does not require exported GCR functions

/**
 * Sets up the Parkruns menu on opening the spreadsheet.
 *  Original instructions to set up (and execute?) the trigger:
 *    1. Go to the Apps Script (Macros) editor.
 *    2. Click on the clock icon (Triggers) in the left sidebar.
 *    3. Click on "Create trigger".
 *    4. Set up the trigger with the following settings:
 *        - Choose function: onOpen
 *        - Select event type: On open
 *        - Save
 */
function onOpen(event) {
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
