using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;

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
        #endregion

        #region Members

        public Sheet? GetSheet(string title, string spreadsheetId)
        {
            return this.sheetsService.Spreadsheets.Get(spreadsheetId).Execute().Sheets.FirstOrDefault(x => x.Properties.Title == title);
        }

        /// <summary>
        /// Cria uma linha nova vazia, no início da tabela
        /// </summary>
        /// <param name="sheetId">ID da aba</param>
        /// <param name="startColumnIndex">Coluna que começa (base 0)</param>
        /// <param name="endColumnIndex">Coluna que termina (base 0)</param>
        /// <param name="spreadsheetId">ID da planilha</param>
        /// <returns></returns>
        public async Task CreateEmptyRowAsync(int? sheetId, int endColumnIndex, string spreadsheetId)
        {
            InsertDimensionRequest insertDimensionRequest = new InsertDimensionRequest
            {
                Range = new DimensionRange
                {
                    Dimension = "ROWS",
                    SheetId = sheetId,
                    StartIndex = 1,
                    EndIndex = 2
                },
                InheritFromBefore = true,
            };

            UpdateCellsRequest updateCellsRequest = new UpdateCellsRequest
            {
                Range = new GridRange
                {
                    SheetId = sheetId,
                    StartRowIndex = 1,
                    EndRowIndex = 2,
                    StartColumnIndex = 0,
                    EndColumnIndex = endColumnIndex,
                },
                Rows = new List<RowData>
                {
                    new RowData
                    {
                        Values = this.CreateEmptyCellDataList(endColumnIndex)
                    }
                },
                Fields = "userEnteredFormat.backgroundColor.red,userEnteredFormat.backgroundColor.green,userEnteredFormat.backgroundColor.red,userEnteredFormat.borders.top,userEnteredFormat.borders.right,userEnteredFormat.borders.bottom,userEnteredFormat.borders.left"
            };

            await this.sheetsService.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
                {
                    new Request { InsertDimension = insertDimensionRequest },
                    new Request { UpdateCells = updateCellsRequest }
                }
            }, spreadsheetId).ExecuteAsync();
        }

        /// <summary>
        /// Cria uma lista de células vazias, com fundo verde, de acordo com a quantidade especificada
        /// </summary>
        /// <param name="count">Quantidade</param>
        /// <returns>List<CellData></returns>
        public List<CellData> CreateEmptyCellDataList(int count)
        {
            List<CellData> values = new List<CellData>();

            for (int i = 0; i < count; i++)
            {
                values.Add(new CellData
                { UserEnteredFormat = new CellFormat
                    {
                        Borders = new Borders {
                            Top = new Border { Style = "SOLID" },
                            Right = new Border { Style = "SOLID" },
                            Bottom = new Border { Style = "SOLID" },
                            Left = new Border { Style = "SOLID" }
                        },
                        BackgroundColor = new Color
                        {
                            Red = 217f / 255f,
                            Green = 234f / 255f,
                            Blue = 211f / 255f
                        } } } );
            }

            return values;
        }

        /// <summary>
        /// Atualiza a linha com os valores
        /// </summary>
        /// <param name="cellDataList">Lista com os valores de cada célula</param>
        /// <param name="sheetId">ID da aba</param>
        /// <param name="spreadsheetId">ID planilha</param>
        /// <returns>Task</returns>
        public async Task updateCellsAsync(List<CellData> cellDataList, string fields, int? sheetId, string spreadsheetId)
        {
            UpdateCellsRequest updateCellsRequest = new UpdateCellsRequest
            {
                Range = new GridRange
                {
                    SheetId = sheetId,
                    StartRowIndex = 1,
                    EndRowIndex = 2,
                    StartColumnIndex = 0,
                    EndColumnIndex = 5
                },
                Rows = new List<RowData>
                {
                    new RowData
                    {
                        Values = cellDataList
                    }
                },
                Fields = fields
            };

            await this.sheetsService.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
                {
                    new Request { UpdateCells = updateCellsRequest }
                }
            }, spreadsheetId).ExecuteAsync();
        }

        /// <summary>
        /// Ordena a planilha pela coluna informada
        /// </summary>
        /// <param name="startColumnIndex">Coluna inicial</param>
        /// <param name="endColumnIndex">Coluna final</param>
        /// <param name="sheetId">ID da aba</param>
        /// <param name="spreadsheetId">ID da planilha</param>
        /// <returns>Task</returns>
        public async Task SortColumnAsync(int startColumnIndex, int endColumnIndex, int? sheetId, string spreadsheetId)
        {
            SortRangeRequest sortRangeRequest = new SortRangeRequest
            {
                Range = new GridRange
                {
                    SheetId = sheetId,
                    StartRowIndex = 1, // Ignorando o cabeçalho
                    EndRowIndex = 1000,
                    StartColumnIndex = startColumnIndex,
                    EndColumnIndex = endColumnIndex
                },
                SortSpecs = new List<SortSpec>
                {
                    new SortSpec { DimensionIndex = startColumnIndex, SortOrder = "ASCENDING" }
                }
            };

            await this.sheetsService.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { new Request { SortRange = sortRangeRequest } }
            }, spreadsheetId).ExecuteAsync();
        }

        /// <summary>
        /// Remove um ou mais linhas com registros ou sem
        /// </summary>
        /// <param name="startRowIndex">Linha inicial</param>
        /// <param name="endRowIndex">Linha final</param>
        /// <param name="sheetId">ID da aba</param>
        /// <param name="spreadsheetId">ID da planilha</param>
        /// <returns>Task</returns>
        public async Task DeleteRowAsync(int startRowIndex, int endRowIndex, int? sheetId, string spreadsheetId)
        {
            DeleteDimensionRequest deleteDimensionRequest = new DeleteDimensionRequest
            {
                Range = new DimensionRange
                {
                    Dimension = "ROWS",
                    SheetId = sheetId,
                    StartIndex = startRowIndex,
                    EndIndex = endRowIndex
                }
            };

            await this.sheetsService.Spreadsheets.BatchUpdate(new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>
                {
                    new Request { DeleteDimension = deleteDimensionRequest }
                }
            }, spreadsheetId).ExecuteAsync();
        }

        public async Task<ValueRange> GetValuesAsync(string namePage, string spreadsheeId)
        {
            return await this.sheetsService.Spreadsheets.Values.Get(spreadsheeId, $"{namePage}!A1:Z1000").ExecuteAsync();
        }

        public async Task<ValueRange> GetValueByRowAsync(int row, string namePage, string spreadsheetId)
        {
            ValueRange header = await this.sheetsService.Spreadsheets.Values.Get(spreadsheetId, $"{namePage}!A1:Z1").ExecuteAsync();

            if (row == 1)
                return header;

            ValueRange value = await this.sheetsService.Spreadsheets.Values.Get(spreadsheetId, $"{namePage}!A{row}:Z{row}").ExecuteAsync();

            value.Values.Add(header.Values[0]);

            return value;
        }

        public Dictionary<string, int> GetHeaders(IList<object> values)
        {
            Dictionary<string, int> headers = new Dictionary<string, int>();

            for (int i = 0; i < values.Count; i++)
            {
                headers[values[i].ToString()] = i;
            }

            return headers;
        }
        #endregion
    }
}
