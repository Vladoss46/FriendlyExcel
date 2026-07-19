using System.Data;

namespace FriendlyExcel.Models
{
    /// <summary>
    /// In-memory workbook: each sheet is a <see cref="DataTable"/>.
    /// </summary>
    public class XLBook
    {
        /// <summary>
        /// Sheets of the workbook, in order.
        /// </summary>
        public List<DataTable> Sheets { get; set; } = [];

        /// <summary>
        /// Creates a book from existing sheets.
        /// </summary>
        public XLBook(IEnumerable<DataTable> sheets)
        {
            Sheets = sheets.ToList();
        }

        /// <summary>
        /// Creates an empty book.
        /// </summary>
        public XLBook()
        {
        }
    }
}
