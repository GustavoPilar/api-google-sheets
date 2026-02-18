using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Model.Requests;
using Google_Sheets_API.Utils.Enums;

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
                "api-sheets.json",
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

        #region Getters
        public SheetsService GetSheetsService()
        {
            return this.sheetsService;
        }

        /// <summary>
        /// Retorna uma aba pelo titulo.
        /// </summary>
        /// <param name="sheets">Lista de abas</param>
        /// <param name="title">Titulo da aba</param>
        /// <returns>Sheet</returns>
        public Sheet? GetSheetByTitle(IList<Sheet> sheets, string title)
        {
            return sheets.FirstOrDefault(x => x.Properties.Title == title);
        }
        #endregion

        #region 'Request' Members

        #region SpreadsheetsResource.GetRequest
        /// <summary>
        /// Retorna a requisição de uma Planilha. Não retorna o Planilha em si, apenas a requisição.
        /// </summary>
        /// <param name="spreeadsheetId">ID da planilha</param>
        /// <returns>SpreadSheet</returns>
        public SpreadsheetsResource.GetRequest GetSpreadsheetRequest(string spreeadsheetId)
        {
            return this.sheetsService.Spreadsheets.Get(spreeadsheetId);
        }
        #endregion

        #region InsertDimensionRequest
        /// <summary>
        /// Retorna uma requisição de inserção de dimensão.
        /// </summary>
        /// <param name="dimensionEnum">Tipo de dimensação - DimensionsEnum</param>
        /// <param name="dimension">Dimensão</param>
        /// <param name="sheetId">ID da aba</param>
        /// <returns>InsertDimensionRequest</returns>
        public InsertDimensionRequest InsertDimensionRequest(DimensionsEnum dimensionEnum, Dimension dimension, int? sheetId)
        {
            return new InsertDimensionRequest()
            {
                Range = new DimensionRange
                {
                    Dimension = dimensionEnum == DimensionsEnum.ROWS ? "ROWS" : "COLUMNS",
                    SheetId = sheetId,
                    StartIndex = dimension.StartIndex,
                    EndIndex = dimension.EndIndex
                },
                InheritFromBefore = dimension.InheritFromBefore
            };
        }
        #endregion


        #region UpdateCellsRequest
        public UpdateCellsRequest UpdateCellsRequest(int? sheetId, Dimension dimension)
        {
            return new UpdateCellsRequest
            {
                Range = new GridRange
                {
                    SheetId = sheetId,
                    StartRowIndex = dimension.StartIndex,
                    StartColumnIndex = 0,
                    EndRowIndex = dimension.EndIndex,
                    EndColumnIndex = 5

                },
                Rows = new List<RowData>
                {
                    new RowData
                    {
                        Values = new List<CellData>
                        {
                            new CellData { UserEnteredFormat = this.SetBackgroundColorDefault() },
                            new CellData { UserEnteredFormat = this.SetBackgroundColorDefault() },
                            new CellData { UserEnteredFormat = this.SetBackgroundColorDefault() },
                            new CellData { UserEnteredFormat = this.SetBackgroundColorDefault() },
                            new CellData { UserEnteredFormat = this.SetBackgroundColorDefault() },
                        }
                    }
                },
                Fields = "userEnteredFormat.backgroundColor.green, userEnteredFormat.backgroundColor.red,userEnteredFormat.backgroundColor.blue"
            };
        }
        #endregion
        #endregion

        #region 'Execute' Members
        public async Task<BatchUpdateSpreadsheetResponse> BatchUpdateExecute(BatchUpdateSpreadsheetRequest requests, string spreadsheetId)
        {
            return await this.sheetsService.Spreadsheets.BatchUpdate(requests, spreadsheetId).ExecuteAsync();
        }
        #endregion

        #region 'Style' Members
        /// <summary>
        /// Retorna a formatação da célula com fundo verde
        /// </summary>
        /// <returns>CellFormat</returns>
        public CellFormat SetBackgroundColorDefault()
        {
            return new CellFormat { BackgroundColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f } };
        }
        #endregion

    }
}
