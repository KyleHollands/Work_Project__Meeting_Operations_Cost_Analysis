function CALCULATE_MOD_COST(dataRange, month, modType) {

    var internalModCost = 0
    var externalModCost = 0
    var allModCost = 0
    var exCount = 0
  
    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
  
    //// New Tiers-------------------------------------------------
    // Rate: $35/h
    var mod_train = [] 
  
    // Rate: $30/h
    var all_mods_new = []
  
    //// Old Tiers-------------------------------------------------
    // Rate: $30/h
    var internal_mods = []
    
    // Rate: <149 Units: $260 | >150 Units: $300 | Associations: $350.
    var external_mods = []
  
    var all_mods = []
  
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      // If sheet matches month
      if (sheet.getName().startsWith(month)) { 
  
        // Current Moderator Pricing Model
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (internal_mods.includes(col) && internal_mods.includes(row[0,0])) {internalModCost += 1 * 30 * 2.5}
            else if (mod_train.includes(col) && mod_train.includes(row[0,0])) {internalModCost += 1 * 35 * 2.5}
            else if (internal_mods.includes(col) && internal_mods.includes(row[0,0]) && row[0,5].startsWith("Orientation"))
            {internalModCost += 1 * 30 * 1}
            else if (mod_train.includes(col) && mod_train.includes(row[0,0]) && row[0,5].startsWith("Orientation"))
            {internalModCost += 1 * 35 * 1}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,5].startsWith("Association")) {externalModCost += 350, exCount += 1}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] <= 149) {externalModCost += 260, exCount += 1}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] >= 150) {externalModCost += 300, exCount += 1}
            else if (external_mods.includes(col) && external_mods.includes(row[0,0]) && row[0,8] == "") {externalModCost += 280, exCount += 1}
            else {}
          });
        });
  
        // Cost if all Moderators were External at Current Price Tiers
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,5].startsWith("Association")) {allModCost += 350}
            else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] <= 149) {allModCost += 260}
            else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] >= 150) {allModCost += 300}
            else if (all_mods.includes(col) && all_mods.includes(row[0,0]) && row[0,8] == "") {allModCost += 280}
            else {}
          });
        });
  
        // Cost of External Moderators at Various Tiers
      } potentialModCostIfExternalInternal = exCount * 30 * 2.5
        externalModCostFifteenPerc = exCount * (30 * 1.15) * 2.5
        externalModCostThirtyPerc = exCount * (30 * 1.30) * 2.5
        externalModCostFortyFivePerc = exCount * (30 * 1.45) * 2.5
        externalModCostSixtyPerc = exCount * (30 * 1.60) * 2.5
        externalModCostOneTwentyPerc = exCount * (30 * 2.20) * 2.5
  
    } // End of Loop
  
    if (modType.startsWith("Internal Mods - Approx. Cost")) {return internalModCost}
    else if (modType.startsWith("Total - Approx. Potential Cost if all External")) {return allModCost}
    else if (modType.startsWith("External Mods - Approx. Cost")) {return externalModCost}
    else if (modType.startsWith("Potential Cost if External Mods at +15%")) {return externalModCostFifteenPerc}
    else if (modType.startsWith("Potential Cost if External Mods at +30%")) {return externalModCostThirtyPerc}
    else if (modType.startsWith("Potential Cost if External Mods at +45%")) {return externalModCostFortyFivePerc}
    else if (modType.startsWith("Potential Cost if External Mods at +60%")) {return externalModCostSixtyPerc}
    else if (modType.startsWith("Potential Cost if External Mods at +120%")) {return externalModCostOneTwentyPerc}
    else if (modType.startsWith("Potential Cost Diff")) {return potentialModCostIfExternalInternal}
    else {}
  } 