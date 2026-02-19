using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Model.Entities;
using Google_Sheets_API.Model.Entities.Base;
using Google_Sheets_API.Services;
using Google_Sheets_API.Services.Base;
using Microsoft.AspNetCore.Mvc;

namespace Google_Sheets_API.Base
{
    public abstract class CrudBaseController<TEntity, TService> : ControllerBase
        where TEntity : class, IEntityBase
        where TService : class, IServiceBase<TEntity>
    {
        #region Fields
        protected readonly GoogleSheetsApiService GoogleSheetsApiService;
        protected readonly string SpreadSheetId = "19LIxqdqwdNPGRnmXhzTDHsO4EQB0Fwf793wdVlJanFU";
        protected readonly TService EntityService;
        #endregion

        #region Constructor
        public CrudBaseController(GoogleSheetsApiService googleSheetsApiService, TService entityService)
        {
            this.GoogleSheetsApiService = googleSheetsApiService;
            this.EntityService = entityService;
        }
        #endregion

        #region 'Get' Actions
        [HttpGet]
        public async Task<ActionResult<IEnumerable<TEntity>>> GetMembers()
        {
            try
            {
                ValueRange values = await this.GoogleSheetsApiService.GetValuesAsync(TEntity.NamePage(), this.SpreadSheetId);

                if (values is null)
                    return NotFound("Nenhum registro encontrado");

                Dictionary<string, int> headers = this.GoogleSheetsApiService.GetHeaders(values.Values[0]);

                List<TEntity> members = new List<TEntity>();

                for (int i = 1; i < values.Values.Count; i++)
                {
                    TEntity member = this.EntityService.TransformValuesToEntity(values.Values[i], headers);
                    member.Id = i + 1;
                    members.Add(member);
                }

                return Ok(members);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.StackTrace);
                return BadRequest($"ERROR => Ocorreu um erro durante da execução: {exception.Message}");
            }
        }

        [HttpGet("{row:int}")]
        public async Task<ActionResult<TEntity>> GetMemberById(int row)
        {
            try
            {
                if (row <= 1)
                    return BadRequest("Id inválido");

                ValueRange values = await this.GoogleSheetsApiService.GetValueByRowAsync(row, TEntity.NamePage(), this.SpreadSheetId);

                if (values is null)
                    return NotFound("Nenhum registro encontrado");

                Dictionary<string, int> headers = this.GoogleSheetsApiService.GetHeaders(values.Values[1]);
                TEntity member = this.EntityService.TransformValuesToEntity(values.Values[0], headers);
                member.Id = row;

                return Ok(member);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.StackTrace);
                return BadRequest($"ERROR => Ocorreu um erro durante a execução: {exception.Message}");
            }
        }
        #endregion

        #region 'Post' Actions
        [HttpPost]
        public async Task<ActionResult<TEntity>> Post(TEntity member)
        {
            try
            {
                Sheet? sheet = this.GoogleSheetsApiService.GetSheet(TEntity.NamePage(), this.SpreadSheetId);

                if (sheet is null)
                    return BadRequest($"A aba {TEntity.NamePage()} não foi encontrada");

                ValueRange valueRange = await this.GoogleSheetsApiService.GetValueByRowAsync(1, sheet.Properties.Title, this.SpreadSheetId);
                Dictionary<string, int> headers = this.GoogleSheetsApiService.GetHeaders(valueRange.Values[0]);

                string fields = "userEnteredValue.stringValue,userEnteredValue.NumberValue,userEnteredFormat.numberFormat,userEnteredFormat.backgroundColor,userEnteredFormat.horizontalAlignment";

                await this.GoogleSheetsApiService.CreateEmptyRowAsync(sheet.Properties.SheetId, headers.Count, this.SpreadSheetId);
                await this.GoogleSheetsApiService.updateCellsAsync(this.EntityService.CreateCellDataList(member), fields, sheet.Properties.SheetId, this.SpreadSheetId);
                await this.GoogleSheetsApiService.SortColumnAsync(0, 5, sheet.Properties.SheetId, this.SpreadSheetId);

                return Ok(member);
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.Message);
                return BadRequest($"ERROR => Ocorreu um erro durante a execução: {exception.Message}");
            }
        }

        #endregion

        #region 'Put' Actions
        #endregion

        #region 'Delete' Actions
        [HttpDelete("DeleteLine")]
        public async Task<ActionResult> DeleteLine(int startRowIndex, int deleteRowIndex)
        {
            try
            {
                Sheet? sheet = this.GoogleSheetsApiService.GetSheet(TEntity.NamePage(), this.SpreadSheetId);

                if (sheet is null)
                    return BadRequest($"a aba {TEntity.NamePage()} não foi encontrada");

                await this.GoogleSheetsApiService.DeleteRowAsync(startRowIndex, deleteRowIndex, sheet.Properties.SheetId, this.SpreadSheetId);

                return Ok($"Linha(s) {startRowIndex} - {deleteRowIndex} foram deletadas");
            }
            catch (Exception exception)
            {
                Console.WriteLine(exception.StackTrace);
                return BadRequest($"ERROR => Ocorreu um erro durante a execução: {exception.Message}");
            }
        }
        #endregion
    }
}
