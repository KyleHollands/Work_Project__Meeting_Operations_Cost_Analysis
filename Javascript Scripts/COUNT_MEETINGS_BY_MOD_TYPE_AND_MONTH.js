function COUNT_MEETINGS_BY_MOD_TYPE_AND_MONTH(dataRange, month, meetingType) {

    var inCount = 0
    var inCountTwo = 0
    var inCountThree = 0
    var inCountFour = 0
    var exCount = 0
    var exCountTwo = 0
    var exCountThree = 0
    var exCountFour = 0
    var allCountOne = 0
    var allCountTwo = 0
    var allCountThree = 0
    var allCountFour = 0
  
    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
  
    //// New Tiers-------------------------------------------------
    var all_mods = []
  
    //// Old Tiers-------------------------------------------------
    //
    //var internal_mods = []
  
    //var external_mods = []
  
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      // If sheet matches month
      if (sheet.getName().startsWith(month)) { 
  
          //// New Code-------------------------------------------------
          values.forEach(function(row) {
          row.forEach(function(col) {
              if (all_mods.includes(col) && row[0,5].startsWith("AB Condo")) {allCountOne += 1}
              else if (all_mods.includes(col) && row[0,5].startsWith("Shareholder")) {allCountFour += 1}
              else if (all_mods.includes(col) && row[0,5].startsWith("Association")) {allCountThree += 1}
              else if (all_mods.includes(col) && !row[0,5].startsWith("Association") && !row[0,5].startsWith("AB Condo") && !row[0,5].startsWith("Orientation") && !row[0,5].startsWith("Shareholder")) {allCountTwo += 1}
              else {}
          });
        });
  
          //// Old Code-------------------------------------------------
          //values.forEach(function(row) {
            //row.forEach(function(col) {
              //if (internal_mods.includes(col) && row[0,5].startsWith("AB Condo")) {inCount += 1}
              //else if (external_mods.includes(col) && row[0,5].startsWith("AB Condo")) {exCount += 1}
              //else if (internal_mods.includes(col) && row[0,5].startsWith("Shareholder")) {inCountFour += 1}
              //else if (external_mods.includes(col) && row[0,5].startsWith("Shareholder")) {exCountFour += 1}
              //else if (internal_mods.includes(col) && row[0,5].startsWith("Association")) {inCountThree += 1}
              //else if (external_mods.includes(col) && row[0,5].startsWith("Association")) {exCountThree += 1}
              //else if (internal_mods.includes(col) && !row[0,5].startsWith("Association") && !row[0,5].startsWith("AB Condo") && !row[0,5].startsWith("Orientation") && !row[0,5].startsWith("Shareholder")) {inCountTwo += 1}
              //else if (external_mods.includes(col) && !row[0,5].startsWith("Association") && !row[0,5].startsWith("AB Condo") && !row[0,5].startsWith("Orientation") && !row[0,5].startsWith("Shareholder")) {exCountTwo += 1}
              //else {}
          //});
        //});
      }
    } // End of Loop
  
    //// New Code-------------------------------------------------
    if (meetingType.startsWith("AB Condo")) {return allCountOne}
    else if (meetingType.startsWith("Association")) {return allCountThree}
    else if (meetingType.startsWith("Ontario") ) {return allCountTwo}
    else if (meetingType.startsWith("Shareholder")) {return allCountFour}
    else {}
  
    //// Old Code-------------------------------------------------
    //if (meetingType.startsWith("AB Condo") && meetingType.startsWith("Internal",28)) {return inCount}
    //else if (meetingType.startsWith("Association") && meetingType.startsWith("Internal", 31)) {return inCountThree}
    //else if (meetingType.startsWith("AB Condo") && meetingType.startsWith("External",28)) {return exCount}
    //else if (meetingType.startsWith("Association") && meetingType.startsWith("External", 31)) {return exCountThree}
    //else if (meetingType.startsWith("Ontario") && meetingType.startsWith("Internal", 27)) {return inCountTwo}
    //else if (meetingType.startsWith("Ontario") && meetingType.startsWith("External", 27)) {return exCountTwo}
    //else if (meetingType.startsWith("Shareholder") && meetingType.startsWith("Internal", 31)) {return inCountFour}
    //else if (meetingType.startsWith("Shareholder") && meetingType.startsWith("External", 31)) {return exCountFour}
    //else if (meetingType.startsWith("Ontario") && !meetingType.startsWith("AB Condo") && !meetingType.startsWith("Association") && !meetingType.startsWith("Shareholder")) {return inCountTwo + exCountTwo}
    //else if (!meetingType.startsWith("Ontario") && !meetingType.startsWith("AB Condo") && !meetingType.startsWith("Shareholder")) {return inCountThree + exCountThree}
    //else if (!meetingType.startsWith("Ontario") && !meetingType.startsWith("Association") && !meetingType.startsWith("Shareholder")) {return inCount + exCount}
    //else if (!meetingType.startsWith("Ontario") && !meetingType.startsWith("AB Condo") && !meetingType.startsWith("Association")) {return inCountFour + exCountFour}
    //else {}
  }