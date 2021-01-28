using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Services;
using System;
using System.Collections.Generic;
using System.IO;

namespace Projecto_Excels
{

    public class GoogleSheetParameters
    {
        public string FromCell { get; set; }
        public string ToCell { get; set; }
        public string SheetName { get; set; }
    }

    public class GoogleDiscourse
    {
        static string[] Scopes = { SheetsService.Scope.Spreadsheets };
        static string ApplicationName = "foroprg";

        private readonly SheetsService _sheetsService;
        private readonly string _spreadsheetId;

        public GoogleDiscourse(string credentialFileName, string spreadsheetId)
        {
            var credential = GoogleCredential.FromStream(new FileStream(credentialFileName, FileMode.Open)).CreateScoped(Scopes);

            _spreadsheetId = spreadsheetId;
            _sheetsService = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        public IList<IList<Object>> GetDataFromSheet(GoogleSheetParameters googleSheetParameters)
        {
            var range = $"{googleSheetParameters.SheetName}!{googleSheetParameters.FromCell}:{googleSheetParameters.ToCell}";
            SpreadsheetsResource.ValuesResource.GetRequest request =
                _sheetsService.Spreadsheets.Values.Get(_spreadsheetId, range);

            var response = request.Execute();
            return response.Values;
        }
    }
}