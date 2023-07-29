function CALCULATE_MOD_COST_BAK(dataRange, month, modType) {

    var internalModCost = 0
    var externalModCost = 0
    var exCount = 0
  
    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
  
    //var exclude_sheets = ["Moderators","MOps OKRs","Meeting Volume"];
    var internal_mods = []
    var external_mods = []
  
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      // If sheet matches month
      if (sheet.getName().startsWith(month)) { 
  
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (internal_mods.includes(col) && internal_mods.includes(row[0,0])) {internalModCost += 1 * 30 * 2.5}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,5].startsWith("Association")) {externalModCost += 350, exCount += 1}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] <= 149) {externalModCost += 260, exCount += 1}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] >= 150) {externalModCost += 300, exCount += 1}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] == "") {externalModCost += 280, exCount += 1}
            else {}
          });
        });
      } potentialModCost = exCount * 30 * 2.5
    } // End of Loop
  
    if (modType.startsWith("Internal")) {return internalModCost}
    else if (modType.startsWith("External")) {return externalModCost}
    else if (modType.startsWith("Potential")) {return potentialModCost}
    else {}
  }