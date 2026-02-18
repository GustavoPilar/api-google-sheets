using Google.Apis.Sheets.v4.Data;

namespace Google_Sheets_API.Model.Requests
{
    public class member<T> where T : class
    {
        public Dimension? dimension { get; set; }
        public T? Entity { get; set; }
    }
}
