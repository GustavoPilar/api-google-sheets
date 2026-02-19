using Google.Apis.Sheets.v4.Data;

namespace Google_Sheets_API.Services.Base
{
    public interface IServiceBase<T> where T : class
    {
        T TransformValuesToEntity(IList<object> values, Dictionary<string, int> headers);
        List<CellData> CreateCellDataList(T entity);
    }
}
