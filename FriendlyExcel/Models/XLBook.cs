using System.Data;

namespace FriendlyExcel.Models
{
    public class XLBook
    {
        List<DataTable> Sheets { get; set; }

        public XLBook(IEnumerable<DataTable> sheets)
        {
            Sheets = sheets.ToList();
        }
    }
}
