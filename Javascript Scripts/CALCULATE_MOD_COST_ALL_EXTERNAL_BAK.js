function CALCULATE_MOD_COST_ALL_EXTERNAL_BAK(dataRange, month, modType) {

    var allModCost = 0
  
    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
  
    //var exclude_sheets = ["Moderators","MOps OKRs","Meeting Volume"];
    var all_mods = []
  
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      // If sheet matches month
      if (sheet.getName().startsWith(month)) { 
  
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,5].startsWith("Association")) {allModCost += 350}
            else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] <= 149) {allModCost += 260}
            else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] >= 150) {allModCost += 300}
            else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] == "") {allModCost += 280}
            else {}
          });
        });
      }
    } // End of Loop
  
    return allModCost
  } 
  