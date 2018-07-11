using System;
using System.Collections.Generic;
using Google.GData.Client;
using Google.GData.Documents;
using Google.GData.Spreadsheets;

namespace M8.SpreadsheetParser {
    public class GoogleDatabase {
        public AtomEntry entry { get; private set; }

        public WorksheetEntry[] worksheets { get; private set; }

        public GoogleDatabase(AtomEntry entry) {
            this.entry = entry;

            RefreshWorksheets();
        }

        private void RefreshWorksheets() {
            //grab worksheet names
            var spreadsheet = this.entry as SpreadsheetEntry;
            var wsFeed = spreadsheet.Worksheets;

            // Iterate through each worksheet in the spreadsheet.
            var worksheetList = new List<WorksheetEntry>();

            foreach(WorksheetEntry worksheetEntry in wsFeed.Entries)
                worksheetList.Add(worksheetEntry);

            worksheets = worksheetList.ToArray();
        }
    }
}