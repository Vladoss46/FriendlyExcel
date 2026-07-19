// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using FriendlyExcel.Models;
using System.Data;

namespace FriendlyExcel.Tests
{
    internal class ExcelReaderShould
    {
        private static readonly string TestDataPath = Path.Combine(AppContext.BaseDirectory, "test-data-sheets");
        private static readonly string[] TestPaths = Directory.GetFiles(TestDataPath);

        [Test]
        public void GetCorrectly([ValueSource(nameof(TestPaths))] string filePath, [Values(0, 1)] int sheetIndex)
        {
            DataTable table = ExcelReader.Get(filePath, sheetIndex, useFirstRowAsColumnNames: true);

            Assert.That(table, Is.Not.Null);
            Assert.That(table.Columns.Count, Is.GreaterThan(0));
            Assert.That(table.Rows.Count, Is.GreaterThan(0));
        }

        private static IEnumerable<TestCaseData> GetBookCases()
        {
            yield return new TestCaseData(TestPaths.Single(x => x.EndsWith(".xls", StringComparison.OrdinalIgnoreCase)), true, false);
            yield return new TestCaseData(TestPaths.Single(x => x.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)), true, true);
        }

        [Test]
        [TestCaseSource(nameof(GetBookCases))]
        public void GetBookCorrectly(string filePath, bool useFirstRowAsColumnNames1, bool useFirstRowAsColumnNames2)
        {
            XLBook book = ExcelReader.GetBook(filePath, useFirstRowAsColumnNames1, useFirstRowAsColumnNames2);

            Assert.That(book, Is.Not.Null);
            Assert.That(book.Sheets.Count, Is.EqualTo(3));
            Assert.That(book.Sheets[0].Columns.Count, Is.GreaterThan(0));
        }

        [Test]
        public void GetBookParamsConnectedCorrectly([ValueSource(nameof(TestPaths))] string filePath)
        {
            XLBook book = ExcelReader.GetBookParamsConnected(filePath, useFirstRowAsColumnNames: true);

            Assert.That(book, Is.Not.Null);
            Assert.That(book.Sheets.Count, Is.EqualTo(3));
            Assert.That(book.Sheets[0].Rows.Count, Is.GreaterThan(0));
        }
    }
}
