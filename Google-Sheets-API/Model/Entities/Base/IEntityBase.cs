namespace Google_Sheets_API.Model.Entities.Base
{
    public interface IEntityBase
    {
        public int Id { get; set; }

        static string GetNamePage() { return ""; }
    }
}
