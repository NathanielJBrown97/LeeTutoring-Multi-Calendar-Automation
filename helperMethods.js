function getTutorName(calendarId) {
    return tutorNameMap[calendarId] || 'Edward';
}

function resetSheets(spreadsheet) {
    var sheets = spreadsheet.getSheets();
    for (var j = 0; j < sheets.length; j++) {
        sheets[j].clear();
    }
}

function getStartOfWeek() {
    var today = new Date();
    var startOfWeek = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    startOfWeek.setDate(today.getDate() - today.getDay());
    return startOfWeek;
}

function calculateWeekStartDate(startOfWeek, weekOffset) {
    var weekStartDate = new Date(startOfWeek);
    weekStartDate.setDate(startOfWeek.getDate() + weekOffset * 7);
    return weekStartDate;
}

function calculateWeekEndDate(weekStartDate) {
    var weekEndDate = new Date(weekStartDate);
    weekEndDate.setDate(weekStartDate.getDate() + 7);
    return weekEndDate;
}

function createOrClearSheet(spreadsheet, sheetName) {
    var sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
        sheet = spreadsheet.insertSheet(sheetName);
    } else {
        sheet.clear();
    }
    return sheet;
}

function setupSheetHeaders(sheet, weekStartDate) {
    var daysOfWeek = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    for (var j = 0; j < 7; j++) {
        var day = new Date(weekStartDate);
        day.setDate(weekStartDate.getDate() + j);
        var dayHeader = daysOfWeek[day.getDay()] + " " + Utilities.formatDate(day, Session.getScriptTimeZone(), "M/d");
        sheet.getRange(1, j + 2).setValue(dayHeader); // Columns B to H
    }

    var startHour = 7;
    var startMinute = 30;
    for (var row = 2; row <= 28; row++) { // Only up to row 28 for 8:30 PM
        var time = new Date(weekStartDate);
        time.setHours(startHour);
        time.setMinutes(startMinute);
        sheet.getRange(row, 1).setValue(Utilities.formatDate(time, Session.getScriptTimeZone(), "h:mm a"));
        startMinute += 30;
        if (startMinute >= 60) {
            startMinute = 0;
            startHour++;
        }
    }
}
function stylizeSheet(sheet, tutorName) {
    // Tutor Label
    sheet.getRange('A1').setValue(tutorName).setFontWeight('bold');
    // x and y headers and bottom bar
    sheet.getRange('A1:A28').setBackground('#B0B0B0'); // Medium Grey
    sheet.getRange('B1:H1').setBackground('#B0B0B0');
    sheet.getRange('A29:H29').setBackground('#B0B0B0');
    sheet.getRange('I1:H29').setBackground('#B0B0B0');
    sheet.getRange('B1:H1').setHorizontalAlignment('center').setFontWeight('bold'); // Make B1 through H1 centered and bold
    sheet.getRange('A2:A28').setHorizontalAlignment('center').setFontWeight('bold'); // Make A2 through A28 centered and bold
    // background of schedule
    sheet.getRange('B2:H28').setBackground('#D3D3D3'); // Light Grey
    // legend outline
    sheet.getRange('K1:M10').setBackground('#B0B0B0'); // Medium Grey
    sheet.getRange('L1').setValue('Legend');
    sheet.getRange('L2').setValue('Manual Update').setBackground('red'); // Take a guess
    sheet.getRange('L3').setValue('Hourly').setBackground('#FFCCCC'); // Light Red
    sheet.getRange('L4').setValue('Recurring').setBackground('#CCFFCC'); // Light Green
    sheet.getRange('L5').setValue('Initial Consult').setBackground('#FFFFCC'); // Light Yellow
    sheet.getRange('L6').setValue('Phone Call').setBackground('#E6E6FA'); // Light Purple
    sheet.getRange('L7').setValue('Training').setBackground('#FFC864'); // Light Orange
    sheet.getRange('L8').setValue('Cancellation').setBackground('#ADD8E6'); // Light Blue
    sheet.getRange('L9').setValue('Filler').setBackground('#D2B48C'); // Light Brown
}

function populateSheetWithEvents(sheet, events, weekStartDate) {
    events.forEach(event => {
        var eventStart = event.getStartTime();
        var eventEnd = event.getEndTime();
        var title = event.getTitle();
        var description = event.getDescription();

        var isFiller = title.toLowerCase().includes('filler');
        var isCancelled = title.toLowerCase().includes('cancelled');
        var isInitialConsult = description.toLowerCase().includes('initial consultation');
        var isHourly = description.toLowerCase().includes('tutoring appointment');
        var isRecurring = description.toLowerCase().includes('recurring');
        var isCall = description.toLowerCase().includes('phone call');

        var studentName = title;

        var dayOffset = calculateDayOffset(eventStart, weekStartDate);
        var { startRow, endRow } = calculateRowRange(eventStart, eventEnd);

        for (var row = startRow; row <= endRow; row++) {
            var range = sheet.getRange(row, dayOffset + 2).setValue(studentName);

            // Highlight cells based on the type of event
            if (isCancelled) {
                range.setBackground('#ADD8E6'); // Light Blue
            } else if (isHourly) {
                range.setBackground('#FFCCCC'); // Light Red
            } else if (isRecurring) {
                range.setBackground('#CCFFCC'); // Light Green
            } else if (isInitialConsult) {
                range.setBackground('#FFFFCC'); // Light Yellow
            } else if (isCall) {
                range.setBackground('#E6E6FA'); // Light Purple
            } else if (isFiller) {
                range.setBackground('#D2B48C'); // Light Brown
            } else {
                range.setBackground('red'); // Take a guess
            }
        }
    });
}

function calculateDayOffset(eventStart, weekStartDate) {
    var eventStartDate = new Date(eventStart.getFullYear(), eventStart.getMonth(), eventStart.getDate());
    return Math.floor((eventStartDate - weekStartDate) / (24 * 60 * 60 * 1000));
}

function calculateRowRange(eventStart, eventEnd) {
    var startHour = eventStart.getHours();
    var startMinute = eventStart.getMinutes();
    var endHour = eventEnd.getHours();
    var endMinute = eventEnd.getMinutes();

    if (startHour < 7 || (startHour === 7 && startMinute < 30)) {
        startHour = 7;
        startMinute = 30;
    }

    if (eventEnd.getDate() !== eventStart.getDate() && endHour === 0 && endMinute === 0) {
        endHour = 20;
        endMinute = 30;
    } else if (endHour > 20 || (endHour === 20 && endMinute > 30)) {
        endHour = 20;
        endMinute = 30;
    }

    var durationMinutes = (endHour * 60 + endMinute) - (startHour * 60 + startMinute);
    var numberOfBlocks = Math.ceil(durationMinutes / 30);

    var startRow = Math.max((startHour - 7) * 2 + (startMinute >= 30 ? 1 : 0) + 1, 2);
    var endRow = startRow + numberOfBlocks - 1;

    if (endRow > 28) {
        endRow = 28;
    }

    return { startRow, endRow };
}

function parseDuration(durationText) {
    var hoursMatch = durationText.match(/(\d+)\s*hour/);
    var minutesMatch = durationText.match(/(\d+)\s*minute/);
    
    var hours = hoursMatch ? parseInt(hoursMatch[1], 10) : 0;
    var minutes = minutesMatch ? parseInt(minutesMatch[1], 10) : 0;
    
    return (hours * 60) + minutes;
  }


  function getBackgroundColorForTutor(tutorName) {
    switch (tutorName) {
      case 'Edward':
        return '#E6F7FF'; // Light Blue
      case 'Eli':
        return '#FFE6E6'; // Light Red
      case 'Kieran':
        return '#E6FFE6'; // Light Green
      case 'Kyra':
        return '#FFCCCB'; // Light Coral
      case 'Patrick':
        return '#F9E6FF'; // Light Purple
      case 'Ben':
        return '#FFF9E6'; // Light Yellow
      default:
        return '#D3D3D3'; // Default Light Grey
    }
  }
  