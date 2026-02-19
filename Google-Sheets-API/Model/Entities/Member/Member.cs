
using Google_Sheets_API.Model.Entities.Base;

namespace Google_Sheets_API.Model.Entities
{
    public class Member : IEntityBase
    {
        public int Id { get; set; }
        public string? Name { get; set; }
        public bool Status { get; set; }
        public DateTime Birthday { get; set; }
        public int Age { get; set; }
        public string? Class { get; set; }

        public static string GetNamePage()
        {
            return "Membros";
        }
    }
}
