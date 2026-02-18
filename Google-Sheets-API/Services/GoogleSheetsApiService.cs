using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Model.Requests;
using Google_Sheets_API.Utils.Enums;
using Microsoft.AspNetCore.Http.HttpResults;
using System.Runtime.CompilerServices;

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

        /// <summary>
        /// Cria uma linha nova vazia, no início da tabela
        /// </summary>
        /// <param name="sheetId">ID da aba</param>
        /// <param name="startColumnIndex">Coluna que começa (base 0)</param>
        /// <param name="endColumnIndex">Coluna que termina (base 0)</param>
        /// <param name="spreadsheetId">ID da planilha</param>
        /// <returns></returns>
        public async Task CreateEmptyRow(int? sheetId, int startColumnIndex, int endColumnIndex, string spreadsheetId)
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
                    StartColumnIndex = startColumnIndex,
                    EndColumnIndex = endColumnIndex,
                },
                Rows = new List<RowData>
                {
                    new RowData
                    {
                        Values = this.CreateEmptyCellDataList(endColumnIndex)
                    }
                },
                Fields = "userEnteredFormat.backgroundColor.red,userEnteredFormat.backgroundColor.green,userEnteredFormat.backgroundColor.red"
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
                values.Add(new CellData { UserEnteredFormat = new CellFormat { BackgroundColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f } } } );
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
        public async Task updateCells(List<CellData> cellDataList, string fields, int? sheetId, string spreadsheetId)
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

        public async Task SortColumn(int startColumnIndex, int endColumnIndex, int? sheetId, string spreadsheetId)
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

        public async Task DeleteRow(int startRowIndex, int endRowIndex, int? sheetId, string spreadsheetId)
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
        #endregion
    }
}
