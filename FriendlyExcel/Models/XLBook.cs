using System.Data;

namespace FriendlyExcel.Models
{
    public class XLBook
    {
        public List<DataTable> Sheets { get; set; } = [];

        public XLBook(IEnumerable<DataTable> sheets)
        {
            Sheets = sheets.ToList();
        }
        public XLBook()
        {
        }
    }
}
