function main() {
    for (var i = 0; i < calendarSheetPairs.length; i++) {
        var pair = calendarSheetPairs[i];
        syncCalendarToSheet(pair.calendarId, pair.spreadsheetId);
    }
    performPostSyncOperations();
}