using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Model;
using Google_Sheets_API.Model.Requests;
using Google_Sheets_API.Services;
using Google_Sheets_API.Utils.Enums;
using Microsoft.AspNetCore.Mvc;

namespace Google_Sheets_API.Controllers
{
    [ApiController]
    [Route("[Controller]")]
    public class MemberController : ControllerBase
    {
        #region Fields
        public readonly SheetsService SheetService;
        public readonly GoogleSheetsApiService GoogleSheetsApiService;
        public readonly string SpreadSheetId = "19LIxqdqwdNPGRnmXhzTDHsO4EQB0Fwf793wdVlJanFU";

        public readonly MemberService memberService;
        #endregion

        #region Constructor
        public MemberController(GoogleSheetsApiService googleSheetsApiService, MemberService memberService)
        {
            this.GoogleSheetsApiService = googleSheetsApiService;
            this.SheetService = this.GoogleSheetsApiService.GetSheetsService();
            this.memberService = memberService;
        }
        #endregion

        #region 'Get' Actions
        [HttpGet]
        public ActionResult GetMember()
        {
            SpreadsheetsResource.GetRequest getRequest = this.SheetService.Spreadsheets.Get(this.SpreadSheetId);
            getRequest.IncludeGridData = true;

            Spreadsheet spreadsheet = getRequest.Execute();

            Sheet? sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1");

            if (sheet is null)
                return NotFound("Nenhuma planilha foi encontrada");

            return Ok(sheet);
        }
        #endregion

        #region 'Post' Actions
        [HttpPost()]
        public async Task<ActionResult<Member>> Post(Member member)
        {
            try
            {
                Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(this.SpreadSheetId).Execute();
                Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;
                string fields = "userEnteredValue.stringValue,userEnteredValue.NumberValue,userEnteredFormat.numberFormat,userEnteredFormat.backgroundColor";

                await this.GoogleSheetsApiService.CreateEmptyRow(sheet.Properties.SheetId, 0, 5, this.SpreadSheetId);
                await this.GoogleSheetsApiService.updateCells(this.memberService.CreateNewMember(member), fields, sheet.Properties.SheetId, this.SpreadSheetId);
                await this.GoogleSheetsApiService.SortColumn(0, 5, sheet.Properties.SheetId, this.SpreadSheetId);

                return Ok(member);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return BadRequest($"ERROR: {exception.Message}");
            }
        }

        #endregion

        #region 'Delete' Actions
        [HttpDelete("DeleteLine")]
        public async Task<ActionResult> DeleteLine(int startRowIndex, int deleteRowIndex)
        {
            Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(this.SpreadSheetId).Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;

            await this.GoogleSheetsApiService.DeleteRow(startRowIndex, deleteRowIndex, sheet.Properties.SheetId, this.SpreadSheetId);

            return Ok($"Linha(s) {startRowIndex} - {deleteRowIndex} deletadas");
        }
        #endregion
    }
}
