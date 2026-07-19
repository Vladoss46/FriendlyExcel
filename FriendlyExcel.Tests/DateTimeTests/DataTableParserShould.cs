using System.Data;

namespace FriendlyExcel.Tests.DateTimeTests
{
    internal class DataTableParserShould
    {
        private static readonly string TestDataPath = ".\\test-data-time";
        private static readonly string[] TestPaths = Directory.GetFiles(TestDataPath);

        [Test]
        public void ReadDateTimeCorrectly([ValueSource(nameof(TestPaths))] string filePath)
        {
            DataTable table = ExcelReader.Get(filePath, sheetIndex: 0, useFirstRowAsColumnNames: true);

            Assert.That(table, Is.Not.Null);
            Assert.That(table.Columns.Count, Is.GreaterThan(0));
            Assert.That(table.Rows.Count, Is.GreaterThan(0));

            bool hasDateLikeColumn = table.Columns
                .Cast<DataColumn>()
                .Any(c => c.DataType == typeof(DateTime)
                    || c.DataType == typeof(DateOnly)
                    || c.DataType == typeof(TimeOnly));

            Assert.That(hasDateLikeColumn, Is.True, "Expected at least one date/time column.");
        }
    }
}
