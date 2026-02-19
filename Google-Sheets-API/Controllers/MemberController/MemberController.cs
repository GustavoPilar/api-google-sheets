using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Model.Entities;
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
        public async Task<ActionResult<IEnumerable<Member>>> GetMembers()
        {
            ValueRange values = await this.GoogleSheetsApiService.GetValuesAsync(Member.GetNamePage(), this.SpreadSheetId);

            if (values is null)
                return NotFound("Nenhum registro encontrado");

            Dictionary<string, int> headers = this.GoogleSheetsApiService.GetHeaders(values.Values[0]);

            List<Member> members = new List<Member>();

            for (int i = 1; i < values.Values.Count; i++)
            {
                Member member = this.memberService.TransformValuesToEntity(values.Values[i], headers);
                member.Id = i + 1;
                members.Add(member);
            }

            return Ok(members);
        }

        [HttpGet("{row:int}")]
        public async Task<ActionResult<Member>> GetMemberById(int row)
        {
            if (row <= 1)
                return BadRequest("Id inválido");

            ValueRange values = await this.GoogleSheetsApiService.GetValueByRowAsync(row, Member.GetNamePage(), this.SpreadSheetId);

            if (values is null)
                return NotFound("Nenhum registro encontrado");

            Dictionary<string, int> headers = this.GoogleSheetsApiService.GetHeaders(values.Values[1]);
            Member member = this.memberService.TransformValuesToEntity(values.Values[0], headers);
            member.Id = row;

            return Ok(member);
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

                await this.GoogleSheetsApiService.CreateEmptyRowAsync(sheet.Properties.SheetId, 0, 5, this.SpreadSheetId);
                await this.GoogleSheetsApiService.updateCellsAsync(this.memberService.CreateNewMember(member), fields, sheet.Properties.SheetId, this.SpreadSheetId);
                await this.GoogleSheetsApiService.SortColumnAsync(0, 5, sheet.Properties.SheetId, this.SpreadSheetId);

                return Ok(member);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return BadRequest($"ERROR: {exception.Message}");
            }
        }

        #endregion

        #region 'Put' Actions
        #endregion

        #region 'Delete' Actions
        [HttpDelete("DeleteLine")]
        public async Task<ActionResult> DeleteLine(int startRowIndex, int deleteRowIndex)
        {
            Spreadsheet spreadsheet = this.SheetService.Spreadsheets.Get(this.SpreadSheetId).Execute();
            Sheet sheet = spreadsheet.Sheets.FirstOrDefault(x => x.Properties.Title == "Página1")!;

            await this.GoogleSheetsApiService.DeleteRowAsync(startRowIndex, deleteRowIndex, sheet.Properties.SheetId, this.SpreadSheetId);

            return Ok($"Linha(s) {startRowIndex} - {deleteRowIndex} deletadas");
        }
        #endregion
    }
}
