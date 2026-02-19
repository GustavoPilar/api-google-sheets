using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Base;
using Google_Sheets_API.Model.Entities;
using Google_Sheets_API.Services;
using Microsoft.AspNetCore.Mvc;

namespace Google_Sheets_API.Controllers
{
    [ApiController]
    [Route("[Controller]")]
    public class MemberController : CrudBaseController<Member, MemberService>
    {
        public MemberController(GoogleSheetsApiService googleSheetsApiService, MemberService memberService) : base (googleSheetsApiService, memberService)
        {
        }
    }
}
