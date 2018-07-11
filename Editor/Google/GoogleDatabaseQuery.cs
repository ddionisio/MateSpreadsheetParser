using System;
using System.IO;
using Google.GData.Client;
using Google.GData.Documents;
using Google.GData.Spreadsheets;

namespace M8.SpreadsheetParser {
    /// <summary>
    /// Use this to grab databases (GoogleDatabase) from Google
    /// </summary>
    public class GoogleDatabaseQuery {
        public IService documentService { get; private set; }
        public IService spreadsheetService { get; private set; }

        public GoogleDatabaseQuery(GoogleSettings settings) {
            var requestFactory = GoogleRequest.RefreshAuthenticate(settings);

            var docService = new DocumentsService("database");
            docService.RequestFactory = requestFactory;

            documentService = docService;

            var ssService = new SpreadsheetsService("database");

            ssService.RequestFactory = requestFactory;
            spreadsheetService = ssService;
        }

        public GoogleDatabase GetDatabase(string name, ref string error) {
            try {
                Google.GData.Spreadsheets.SpreadsheetQuery query = new Google.GData.Spreadsheets.SpreadsheetQuery();

                // Make a request to the API and get all spreadsheets.
                SpreadsheetsService service = spreadsheetService as SpreadsheetsService;

                SpreadsheetFeed feed = service.Query(query);

                if(feed.Entries.Count == 0) {
                    error = "There are no spreadsheets found.";
                    return null;
                }

                AtomEntry spreadsheet = null;
                foreach(AtomEntry sf in feed.Entries) {
                    if(sf.Title.Text == name)
                        spreadsheet = sf;
                }

                if(spreadsheet == null) {
                    error = string.Format("Spreadsheet: \"{0}\" is not found.", name);
                    return null;
                }

                return new GoogleDatabase(spreadsheet);
            }
            catch(Exception e) {
                error = e.Message;
                return null;
            }
        }
    }
}