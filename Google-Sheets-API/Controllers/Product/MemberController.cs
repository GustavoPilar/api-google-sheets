using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Model;
using Google_Sheets_API.Services;
using Microsoft.AspNetCore.Mvc;

namespace Google_Sheets_API.Controllers
{
    [ApiController]
    [Route("[Controller]")]
    public class MemberController : ControllerBase
    {
        #region Fields
        public readonly SheetsService SheetService;
        public readonly string SpreadSheetId = "19LIxqdqwdNPGRnmXhzTDHsO4EQB0Fwf793wdVlJanFU";
        #endregion

        #region Constructor
        public MemberController(GoogleSheetsApiService googleSheetsApiService)
        {
            this.SheetService = googleSheetsApiService.GetSheetsService();
        }
        #endregion

        #region 'Get' Actions
        [HttpGet]
        public ActionResult GetProduct()
        {
            SpreadsheetsResource.GetRequest getRequest = this.SheetService.Spreadsheets.Get(this.SpreadSheetId);
            getRequest.IncludeGridData = true;

            Spreadsheet spreadsheet = getRequest.Execute();

            Sheet? sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1");

            if (sheet is null)
                return NotFound("Nenhuma planilha foi encontrada");

            return Ok(sheet.Data[0].RowData);
        }
        #endregion

        #region 'Post' Actions
        [HttpPost("CreateLine")]
        public ActionResult CreateNewLine()
        {
            Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(SpreadSheetId).Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;

            InsertDimensionRequest insertDimensionRequest = new InsertDimensionRequest()
            {
                Range = new DimensionRange
                {
                    Dimension = "ROWS",
                    SheetId = sheet.Properties.SheetId,
                    StartIndex = 1,
                    EndIndex = 2
                },
                InheritFromBefore = false
            };

            var batchRequests = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
                {
                    new Request()
                    {
                        InsertDimension = insertDimensionRequest,
                    }
                }
            };

            var batchUpdateSpreadsheetRequest = this.SheetService.Spreadsheets.BatchUpdate(batchRequests, SpreadSheetId).Execute();

            return Ok("Linha adicionada");
        }

        [HttpPost("CreateColumn")]
        public ActionResult CreateNewColumn()
        {
            Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(this.SpreadSheetId).Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;

            InsertDimensionRequest insertDimensionRequest = new InsertDimensionRequest
            {
                Range = new DimensionRange
                {
                    Dimension = "COLUMNS",
                    SheetId = sheet.Properties.SheetId,
                    StartIndex = 0,
                    EndIndex = 1
                }
            };

            Request request = new Request
            {
                InsertDimension = insertDimensionRequest
            };

            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { request }
            };

            this.SheetService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, this.SpreadSheetId).Execute();

            return Ok("Coluna craida");
        }

        [HttpPost("CreateValue")]
        public ActionResult<Member> PostProduct(int startRowIndex, int endRowIndex, Member member)
        {
            if (member is null)
                return BadRequest("Entidade inválida");

            Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(this.SpreadSheetId).Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;
            var baseDate = new DateTime(1899, 12, 30);

            UpdateCellsRequest updateCellsRequest = new UpdateCellsRequest
            {
                Range = new GridRange
                {
                    SheetId = sheet.Properties.SheetId,
                    StartRowIndex = startRowIndex,
                    EndRowIndex = endRowIndex,
                    StartColumnIndex = 0,
                    EndColumnIndex = 5,
                },
                Rows = new List<RowData>()
                {
                    new RowData
                    {
                        Values = new List<CellData>()
                        {
                            new CellData { UserEnteredValue = new ExtendedValue { StringValue = member.Name }, UserEnteredFormat = new CellFormat { BackgroundColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f } } },
                            new CellData { UserEnteredValue = new ExtendedValue { StringValue = member.Status == true ? "ATIVO" : "INATIVO" }, UserEnteredFormat = new CellFormat { BackgroundColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f } } },
                            new CellData { UserEnteredValue = new ExtendedValue { NumberValue = (member.Birthday - baseDate).TotalDays }, UserEnteredFormat = new CellFormat { BackgroundColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f } } },
                            new CellData { UserEnteredValue = new ExtendedValue { NumberValue = member.Age }, UserEnteredFormat = new CellFormat { BackgroundColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f } } },
                            new CellData { UserEnteredValue = new ExtendedValue { StringValue = member.Class }, UserEnteredFormat = new CellFormat { BackgroundColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f } } },
                        }
                    }
                },
                Fields = "userEnteredValue,userEnteredFormat.backgroundColor.green,userEnteredFormat.backgroundColor.red,userEnteredFormat.backgroundColor.blue"
            };

            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request>()
                {
                    new Request { UpdateCells = updateCellsRequest }
                }
            };

            this.SheetService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, this.SpreadSheetId).Execute();

            return Ok(member);
        }
        #endregion

        #region 'Put' Actions

        #endregion

        #region 'Delete' Actions
        [HttpDelete("DeleteLine")]
        public ActionResult DeleteLine()
        {
            Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(this.SpreadSheetId).Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;

            DeleteDimensionRequest deleteDimensionRequest = new DeleteDimensionRequest
            {
                Range = new DimensionRange
                {
                    Dimension = "ROWS",
                    SheetId = sheet.Properties.SheetId,
                    StartIndex = 1,
                    EndIndex = 2
                }
            };

            List<Request> requests = new List<Request>()
            {
                new Request()
                {
                    DeleteDimension = deleteDimensionRequest
                }
            };

            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest()
            {
                Requests = requests
            };

            this.SheetService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, this.SpreadSheetId).Execute();

            return Ok("Linha deletada");
        }

        [HttpDelete("DeleteColumn")]
        public ActionResult DeleteColumn()
        {
            Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(this.SpreadSheetId).Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;

            DeleteDimensionRequest deleteDimension = new DeleteDimensionRequest
            {
                Range = new DimensionRange
                {
                    Dimension = "COLUMNS",
                    SheetId = sheet.Properties.SheetId,
                    StartIndex = 0,
                    EndIndex = 1
                }
            };

            Request request = new Request
            {
                DeleteDimension = deleteDimension
            };

            BatchUpdateSpreadsheetRequest batchUpdateSpreadsheetRequest = new BatchUpdateSpreadsheetRequest
            {
                Requests = new List<Request> { request }
            };

            this.SheetService.Spreadsheets.BatchUpdate(batchUpdateSpreadsheetRequest, this.SpreadSheetId).Execute();

            return Ok("Coluna deletada");
        }
        #endregion
    }
}
