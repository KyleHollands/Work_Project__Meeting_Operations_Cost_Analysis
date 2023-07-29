function COUNT_MEETINGS_BY_SUPPORT_AND_MONTH(dataRange, month, dataRangeTwo, dataRangeThree) {

    var supportOne = 0
    var supportTwo = 0
    var supportThree = 0
    var supportFour = 0
    var supportFive = 0
    var meetingCount = 0
  
    var ss = SpreadsheetApp.getActive();
    var allsheets = ss.getSheets();
  
    // Acquire number of support per month
    for(var s in allsheets){
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRange)
      var values = range.getValues()
  
      // If sheet matches month
      if (sheet.getName().startsWith(month)) { 
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (col != "" && col.startsWith()) {
              supportOne += 1}
            else if (col != "" && col.startsWith()) {
              supportTwo += 1}
            else if (col != "" && col.startsWith()) {
              supportThree += 1}
            else if (col != "" && col.startsWith()) {
              supportFour += 1}
            else if (col != "" && col.startsWith()) {
              supportFive += 1}
            else if (col != "" && col.startsWith()) {
              supportSix += 1}
            else if (col != "" && col.startsWith()) {
              supportSeven += 1}
            else if (col != "" && col.startsWith()) {
              supportEight += 1}
          });
        });
      } 
    } 
  
    // Acquire total meetings per month
    for(var s in allsheets) {
      var sheet = allsheets[s]
      var range = sheet.getRange(dataRangeTwo)
      var values = range.getValues()
  
      if (sheet.getName().startsWith(month)) { // If sheet matches month
        values.forEach(function(row) {
          row.forEach(function(col) {
            if (col === "") {} else {meetingCount += 1}
          })
        })
      }
    } 
    
    supportTotal = supportOne + supportTwo + supportThree + supportFour + supportFive + supportSix + supportSeven + supportEight
    
    if (dataRangeThree.startsWith("Average")) {return meetingCount / supportTotal}
    else if (dataRangeThree.startsWith("#")) {return supportTotal}
  }
  