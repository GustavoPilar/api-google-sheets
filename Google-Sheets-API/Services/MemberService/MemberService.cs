using Google.Apis.Sheets.v4.Data;
using Google_Sheets_API.Model.Entities;
using Google_Sheets_API.Services.Base;

namespace Google_Sheets_API.Services
{
    public class MemberService : IServiceBase<Member>
    {
        #region Members
        public Member TransformValuesToEntity(IList<object> values, Dictionary<string, int> headers)
        {
            Member member = new Member()
            {
                Name = values[headers["Nome"]]?.ToString(),
                Status = bool.TryParse(values[headers["Status"]]?.ToString(), out var status) && status,
                Birthday = DateTime.TryParse(values[headers["Data de nascimento"]].ToString(), out var date) ? date : DateTime.Now,
                Age = int.TryParse(values[headers["Idade"]].ToString(),out var age) ? age : 0,
                Class = values[headers["Classe"]].ToString()
            };

            return member;
        }


        public List<CellData> CreateCellDataList(Member member)
        {

            List<CellData> list = new List<CellData>
            {
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = member.Name } },
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = member.Status == true ? "ATIVO" : "INATIVO" } },
                new CellData { UserEnteredValue = new ExtendedValue { NumberValue = (member.Birthday - new DateTime(1899, 12, 30)).TotalDays }, UserEnteredFormat = new CellFormat { NumberFormat = new NumberFormat { Type = "DATE", Pattern = "dd/MM/yyyy" }, HorizontalAlignment = "RIGHT" } },
                new CellData { UserEnteredValue = new ExtendedValue { NumberValue = member.Age }, UserEnteredFormat = new CellFormat { HorizontalAlignment = "RIGHT" } },
                new CellData { UserEnteredValue = new ExtendedValue { StringValue = member.Class } }
            };

            foreach (CellData cell in list)
            {
                Color redColor = new Color { Red = 244f / 255f, Green = 204f / 255, Blue = 204f / 255f };
                Color greenColor = new Color { Red = 217f / 255f, Green = 234f / 255f, Blue = 211f / 255f };

                if (cell.UserEnteredFormat is null)
                {
                    cell.UserEnteredFormat = new CellFormat { BackgroundColor = member.Status ? greenColor : redColor };
                }
                else
                {
                    cell.UserEnteredFormat.BackgroundColor = member.Status ? greenColor : redColor;
                }
            }

            return list;
        }
        #endregion
    }
}
