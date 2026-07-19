// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using NPOI.SS.UserModel;
using ExcelDataTableReader = FriendlyExcel.FunctionalClasses.DataTableReader;

namespace FriendlyExcel.Tests.FunctionalClasses
{
    internal class DataTableReaderShould
    {
        private static readonly string TestDataPath = Path.Combine(AppContext.BaseDirectory, "test-data");
        private static readonly string[] TestPaths = Directory.GetFiles(TestDataPath);

        [Test]
        [TestCaseSource(nameof(TestPaths))]
        public void ReadCorrectly(string filePath)
        {
            using IWorkbook book = ExcelDataTableReader.Read(filePath);
            Assert.That(book, Is.Not.Null);
            Assert.That(book.NumberOfSheets, Is.GreaterThan(0));
        }
    }
}
