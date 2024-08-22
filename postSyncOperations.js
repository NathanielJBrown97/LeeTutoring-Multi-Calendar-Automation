function performPostSyncOperations() {
    copyToMasterSheet(); // Syncs Individual Sheets To Master.
    checkForRecurringAppointments(); // Call the recurring check after syncing the sheets
    listStudentsAndDetails(); // Outputs side table showcasing all students booked for the week; related tutors, and time booked.
    colorCodeStudentsAndDetailsTable(); // color codes side table after creation
    sortSheetsByDate(); // Ensure the sheets at bottom left are in numerical order
}



function copyToMasterSheet(masterSpreadsheetId, spreadsheetPairs, ) {
  
    var masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
  
    // Determine the current week
    var today = new Date();
    var startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
  
    // Generate sheets for the past two weeks, current week, and next three weeks
    for (var weekOffset = -2; weekOffset <= 3; weekOffset++) {
      var weekStartDate = new Date(startOfWeek);
      weekStartDate.setDate(startOfWeek.getDate() + weekOffset * 7);
      var sheetName = Utilities.formatDate(weekStartDate, Session.getScriptTimeZone(), "M/d");
  
      var masterSheet = masterSpreadsheet.getSheetByName(sheetName);
      if (!masterSheet) {
        masterSheet = masterSpreadsheet.insertSheet(sheetName);
      } else {
        masterSheet.clear();
      }
  
      // Define starting row and column for each copy-paste block
      var startPositions = [
        {row: 1, col: 1},
        {row: 1, col: 13},
        {row: 32, col: 1},
        {row: 32, col: 13},
        {row: 63, col: 1},
        {row: 63, col: 13}
      ];
  
      for (var i = 0; i < spreadsheetPairs.length; i++) {
        var pair = spreadsheetPairs[i];
        var spreadsheetId = pair.spreadsheetId;
        var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        var sheet = spreadsheet.getSheetByName(sheetName);
  
        if (sheet) {
          var startPos = startPositions[i];
          var range = sheet.getRange('A1:I29');
          var values = range.getValues();
          var backgroundColors = range.getBackgrounds();
          var fonts = range.getFontWeights();
          var fontColors = range.getFontColors();
  
          // Convert time and date values to text to prevent misinterpretation
          for (var row = 0; row < values.length; row++) {
            for (var col = 0; col < values[row].length; col++) {
              if (values[row][col] instanceof Date) {
                // For the first row (header), format as date, for other rows format as time
                if (row === 0) {
                  values[row][col] = Utilities.formatDate(values[row][col], Session.getScriptTimeZone(), "EEE M/d");
                } else {
                  values[row][col] = Utilities.formatDate(values[row][col], Session.getScriptTimeZone(), "h:mm a");
                }
              }
            }
          }
  
          var targetRange = masterSheet.getRange(startPos.row, startPos.col, 29, 9);
          targetRange.setValues(values);
          targetRange.setBackgrounds(backgroundColors);
          targetRange.setFontWeights(fonts);
          targetRange.setFontColors(fontColors);
        } else {
          Logger.log(`Sheet ${sheetName} not found in spreadsheet ${spreadsheetId}`);
        }
      }
  
      // Add "Last Updated" timestamp in J1 and K1
      var lastUpdated = new Date();
      
      masterSheet.getRange('J1').setValue('Last Updated:').setFontWeight('bold').setBackground('#FFA500');
      masterSheet.getRange('K1').setValue(Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), "MMMM dd, yyyy")).setFontWeight('bold').setBackground('#FFA500'); // Full date in K1
      masterSheet.getRange('L1').setValue(Utilities.formatDate(lastUpdated, Session.getScriptTimeZone(), "h:mm a")).setFontWeight('bold').setBackground('#FFA500'); // Time in L1
  
  
      // Shift the first legend down by 1 row and add headers
      masterSheet.getRange('J2:L11').setBackground('#808080'); // Dark Grey for the first legend background
      masterSheet.getRange('K2').setValue('Legend').setFontWeight('bold').setBackground('#B0B0B0');
      masterSheet.getRange('K3').setValue('Manual Update').setBackground('red'); // Manual Update (Red)
      masterSheet.getRange('K4').setValue('Hourly').setBackground('#FFCCCC'); // Hourly (Light Red)
      masterSheet.getRange('K5').setValue('Recurring').setBackground('#CCFFCC'); // Recurring (Light Green)
      masterSheet.getRange('K6').setValue('Initial Consult').setBackground('#FFFFCC'); // Initial Consult (Light Yellow)
      masterSheet.getRange('K7').setValue('Phone Call').setBackground('#E6E6FA'); // Phone Call (Light Purple)
      masterSheet.getRange('K8').setValue('Training').setBackground('#FFC864'); // Training (Light Orange)
      masterSheet.getRange('K9').setValue('Cancellation').setBackground('#ADD8E6'); // Cancellation (Light Blue)
      masterSheet.getRange('K10').setValue('Filler').setBackground('#D2B48C'); // Filler (Light Brown)
  
      // Block Off White Space
      masterSheet.getRange('J12:L29').setBackground('#808080'); // Range J12 through L29
      masterSheet.getRange('A30:I31').setBackground('#808080'); // Range A30 through I31
      masterSheet.getRange('M30:U31').setBackground('#808080'); // Range M30 through U31
      masterSheet.getRange('J39:L60').setBackground('#808080'); // Range J39 through L60
      masterSheet.getRange('A61:I62').setBackground('#808080'); // Range A61 through I62
      masterSheet.getRange('M61:U62').setBackground('#808080'); // Range M61 through U62
      masterSheet.getRange('J70:L91').setBackground('#808080'); // Range J70 through L91
      masterSheet.getRange('V1:V91').setBackground('#808080');  // Range V1 through V91
      masterSheet.getRange('Z1:Z91').setBackground('#808080');  // Range Z1 through Z91
  
  
      // Repeat the legend in line with each tutor's block
      var legendPositions = [
        {row: 30, col: 10},
        {row: 61, col: 10}
      ];
  
      legendPositions.forEach(function(pos) {
        var legendRange = masterSheet.getRange(pos.row, pos.col, 9, 3);
        legendRange.setBackground('#808080'); // Dark Grey for the legend background
        legendRange.getCell(1, 2).setValue('Legend').setFontWeight('bold').setBackground('#B0B0B0');
        legendRange.getCell(2, 2).setValue('Manual Update').setBackground('red'); // Manual Update (Red)
        legendRange.getCell(3, 2).setValue('Hourly').setBackground('#FFCCCC'); // Hourly (Light Red)
        legendRange.getCell(4, 2).setValue('Recurring').setBackground('#CCFFCC'); // Recurring (Light Green)
        legendRange.getCell(5, 2).setValue('Initial Consult').setBackground('#FFFFCC'); // Initial Consult (Light Yellow)
        legendRange.getCell(6, 2).setValue('Phone Call').setBackground('#E6E6FA'); // Phone Call (Light Purple)
        legendRange.getCell(7, 2).setValue('Training').setBackground('#FFC864'); // Training (Light Orange)
        legendRange.getCell(8, 2).setValue('Cancellation').setBackground('#ADD8E6'); // Cancellation (Light Blue)
        legendRange.getCell(9, 2).setValue('Filler').setBackground('#D2B48C'); // Filler (Light Brown)
      });
    }
  }
  
  function checkForRecurringAppointments(masterSpreadsheetId) {
    var masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
  
    // Determine the current week and past/future weeks
    var today = new Date();
    var startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
  
    for (var weekOffset = -2; weekOffset < 3; weekOffset++) {
      var currentWeekStartDate = new Date(startOfWeek);
      currentWeekStartDate.setDate(startOfWeek.getDate() + weekOffset * 7);
      var currentSheetName = Utilities.formatDate(currentWeekStartDate, Session.getScriptTimeZone(), "M/d");
      var currentSheet = masterSpreadsheet.getSheetByName(currentSheetName);
  
      var nextWeekStartDate = new Date(startOfWeek);
      nextWeekStartDate.setDate(startOfWeek.getDate() + (weekOffset + 1) * 7);
      var nextSheetName = Utilities.formatDate(nextWeekStartDate, Session.getScriptTimeZone(), "M/d");
      var nextSheet = masterSpreadsheet.getSheetByName(nextSheetName);
  
      if (!currentSheet || !nextSheet) {
        Logger.log("Sheet not found for one of the weeks. Skipping.");
        continue;
      }
  
      var range = currentSheet.getRange('B2:H28'); // Range of time slots in the current week
      var nextRange = nextSheet.getRange('B2:H28'); // Same range in the next week
      var currentValues = range.getValues();
      var nextValues = nextRange.getValues();
      var currentBackgrounds = nextRange.getBackgrounds();
  
      for (var row = 0; row < currentValues.length; row++) {
        for (var col = 0; col < currentValues[row].length; col++) {
          var currentStudent = currentValues[row][col];
          var nextStudent = nextValues[row][col];
  
          if (currentStudent && nextStudent && currentStudent === nextStudent && currentStudent.toLowerCase().indexOf('filler') === -1) {
            Logger.log(`Marking ${currentStudent} as Recurring at ${row + 2}, ${col + 2}`);
            nextRange.getCell(row + 1, col + 1).setBackground('#CCFFCC'); // Light Green for recurring
          }
        }
      }
    }
  }
  
  function listStudentsAndDetails(masterSpreadsheetId) {
    var masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
  
    // Determine the current week and past/future weeks
    var today = new Date();
    var startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
  
    for (var weekOffset = -2; weekOffset <= 3; weekOffset++) {
      var weekStartDate = new Date(startOfWeek);
      weekStartDate.setDate(startOfWeek.getDate() + weekOffset * 7);
      var sheetName = Utilities.formatDate(weekStartDate, Session.getScriptTimeZone(), "M/d");
      var sheet = masterSpreadsheet.getSheetByName(sheetName);
  
      if (!sheet) {
        Logger.log(`Sheet ${sheetName} not found.`);
        continue;
      }
  
      var tutors = [
        {name: 'Edward', range: 'B2:H28'},
        {name: 'Eli', range: 'N2:T28'},
        {name: 'Kieran', range: 'B33:H59'},
        {name: 'Kyra', range: 'N33:T59'},
        {name: 'Patrick', range: 'B64:H90'},
        {name: 'Ben', range: 'N64:T90'}
      ];
  
      var students = {};
  
      tutors.forEach(function(tutor) {
        var range = sheet.getRange(tutor.range);
        var values = range.getValues();
        var backgrounds = range.getBackgrounds();
  
        for (var row = 0; row < values.length; row++) {
          for (var col = 0; col < values[row].length; col++) {
            var studentName = values[row][col];
            var backgroundColor = backgrounds[row][col];
  
            // Check if the student name exists and the background color is either #FFCCCC (Light Red) or #CCFFCC (Light Green)
            if (studentName && (backgroundColor === '#ffcccc' || backgroundColor === '#ccffcc')) {
              if (!students[studentName]) {
                students[studentName] = {
                  tutors: new Set(),
                  totalMinutes: 0
                };
              }
              students[studentName].tutors.add(tutor.name);
              students[studentName].totalMinutes += 30; // Assuming each block represents 30 minutes
            }
          }
        }
      });
  
      // Clear previous data in the columns W, X, Y
      var colStudent = 23; // Column W
      var colTutor = 24; // Column X
      var colDuration = 25; // Column Y
      sheet.getRange(2, colStudent, sheet.getMaxRows(), 3).clear();
  
      // Add Headers
      sheet.getRange(1, colStudent).setValue("Student").setFontWeight('bold').setBackground('#B0B0B0');
      sheet.getRange(1, colTutor).setValue("Tutors").setFontWeight('bold').setBackground('#B0B0B0');
      sheet.getRange(1, colDuration).setValue("Time Booked This Week").setFontWeight('bold').setBackground('#B0B0B0');
  
      // Write the results to the right of the calendar, starting from row 2
      var startRow = 2; // Start writing from the second row in column W
  
      for (var studentName in students) {
        var tutorList = Array.from(students[studentName].tutors).join(', ');
        var duration = students[studentName].totalMinutes;
  
        // Convert totalMinutes to a more readable format
        var durationText = (duration >= 60 ? Math.floor(duration / 60) + " hour" + (Math.floor(duration / 60) > 1 ? "s " : " ") : "") + (duration % 60 > 0 ? duration % 60 + " minutes" : "").trim();
  
        // Determine the background color based on the first tutor's name
        var backgroundColor = getBackgroundColorForTutor(tutorList.split(', ')[0]);
  
        sheet.getRange(startRow, colStudent).setValue(studentName).setBackground(backgroundColor);
        sheet.getRange(startRow, colTutor).setValue(tutorList).setBackground(backgroundColor);
        sheet.getRange(startRow, colDuration).setValue(durationText).setBackground(backgroundColor);
  
        startRow++;
      }
  
      // Adjust column widths for better readability
      sheet.setColumnWidth(colStudent, 160);
      sheet.setColumnWidth(colTutor, 200);
      sheet.setColumnWidth(colDuration, 160);
  
      sheet.getRange('A92:Z93').setBackground('#808080'); // Range W91 through Y91
    }
  
  }
  
  

  // Helper function to color code the students table for manual attention
  function colorCodeStudentsAndDetailsTable(masterSpreadsheetId) {
    var masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
  
    // Determine the current week and past/future weeks
    var today = new Date();
    var startOfWeek = new Date(today);
    startOfWeek.setDate(today.getDate() - today.getDay());
  
    for (var weekOffset = -2; weekOffset <= 3; weekOffset++) {
      var weekStartDate = new Date(startOfWeek);
      weekStartDate.setDate(startOfWeek.getDate() + weekOffset * 7);
      var sheetName = Utilities.formatDate(weekStartDate, Session.getScriptTimeZone(), "M/d");
      var sheet = masterSpreadsheet.getSheetByName(sheetName);
  
      if (!sheet) {
        Logger.log(`Sheet ${sheetName} not found.`);
        continue;
      }
  
      var colDuration = 25; // Column Y (Time Booked This Week)
      var colStudent = 23; // Column W (Student Name)
      var numRows = sheet.getLastRow();
  
      for (var row = 2; row <= numRows; row++) {
        var durationText = sheet.getRange(row, colDuration).getValue();
        var studentName = sheet.getRange(row, colStudent).getValue();
  
        if (!durationText || !studentName) {
          // If either the duration or student name cell is empty, skip coloring
          continue;
        }
  
        var durationMinutes = parseDuration(durationText);
  
        if (durationMinutes >= 120) {
          // If 2 hours or more, color GREEN
          sheet.getRange(row, colDuration).setBackground('green');
          sheet.getRange(row, colStudent).setBackground('green');
        } else if (durationMinutes >= 30 && durationMinutes <= 90) {
          // If between 30 minutes and 1 hour 30 minutes, color YELLOW
          sheet.getRange(row, colDuration).setBackground('yellow');
          sheet.getRange(row, colStudent).setBackground('yellow');
        } else if (durationMinutes <= 15) {
          // If 15 minutes or less, color RED
          sheet.getRange(row, colDuration).setBackground('red');
          sheet.getRange(row, colStudent).setBackground('red');
        }
      }
    }
  }
  
 
  
  function sortSheetsByDate(masterSpreadsheetId) {
    var masterSpreadsheet = SpreadsheetApp.openById(masterSpreadsheetId);
    var sheets = masterSpreadsheet.getSheets();
  
    // Extract the sheet names and parse them into dates
    var sheetDatePairs = sheets.map(function(sheet) {
      var sheetName = sheet.getName();
      
      // Ensure the sheetName format is MM/DD or M/D
      var dateParts = sheetName.split('/');
      if (dateParts.length !== 2) {
        Logger.log(`Invalid sheet name format: ${sheetName}`);
        return null; // Skip invalid formats
      }
  
      // Parse the date assuming MM/DD format
      var date = new Date();
      date.setMonth(parseInt(dateParts[0], 10) - 1); // Set month (0-based)
      date.setDate(parseInt(dateParts[1], 10)); // Set day
  
      // Return the sheet and corresponding date
      return { sheet: sheet, date: date, sheetName: sheetName };
    }).filter(function(pair) {
      return pair !== null; // Filter out invalid entries
    });
  
    // Sort the sheets based on the parsed dates
    sheetDatePairs.sort(function(a, b) {
      return a.date - b.date;
    });
  
    // Reorder the sheets in the spreadsheet
    for (var i = 0; i < sheetDatePairs.length; i++) {
      masterSpreadsheet.setActiveSheet(sheetDatePairs[i].sheet);
      masterSpreadsheet.moveActiveSheet(i + 1);
    }
  
    Logger.log('Sheets successfully sorted by date.');
  }