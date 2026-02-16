using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;

namespace Google_Sheets_API.Services
{
    public class GoogleSheetsApiService
    {
        #region Fields
        public readonly SheetsService sheetsService;
        #endregion

        #region Constructor
        public GoogleSheetsApiService()
        {
            GoogleCredential credential;

            using (FileStream stream = new FileStream(
                "credentials.josn",
                FileMode.Open,
                FileAccess.Read))
            {
                credential = GoogleCredential.FromStream(stream).CreateScoped(SheetsService.Scope.Spreadsheets);
            }

            this.sheetsService = new SheetsService(
                new BaseClientService.Initializer()
                {
                    HttpClientInitializer = credential,
                    ApplicationName = "Google Sheets API"
                });
        }
        #endregion

        #region Getter
        public SheetsService GetSheetsService()
        {
            return this.sheetsService;
        }
        #endregion

    }
}
