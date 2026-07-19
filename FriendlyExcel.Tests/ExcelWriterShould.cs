// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using FriendlyExcel.Models;
using System.Data;

namespace FriendlyExcel.Tests
{
    internal class ExcelWriterShould
    {
        private string _tempDir = null!;

        [SetUp]
        public void Setup()
        {
            _tempDir = Path.Combine(Path.GetTempPath(), "FriendlyExcelTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
        }

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, recursive: true);
        }

        [TestCase(".xlsx")]
        [TestCase(".xls")]
        public void SaveAndReadBackCorrectly(string extension)
        {
            DataTable source = CreateSampleTable("People");
            string filePath = Path.Combine(_tempDir, $"sample{extension}");

            ExcelWriter.Save(filePath, source);
            Assert.That(File.Exists(filePath), Is.True);

            DataTable loaded = ExcelReader.Get(filePath);

            Assert.That(loaded.Columns.Count, Is.EqualTo(source.Columns.Count));
            Assert.That(loaded.Rows.Count, Is.EqualTo(source.Rows.Count));
            Assert.That(loaded.Columns[0].ColumnName, Is.EqualTo("Name"));
            Assert.That(loaded.Rows[0]["Name"], Is.EqualTo("Alice"));
            Assert.That(Convert.ToInt32(loaded.Rows[0]["Age"]), Is.EqualTo(30));
            Assert.That(Convert.ToDouble(loaded.Rows[0]["Score"]), Is.EqualTo(12.5).Within(0.001));
            Assert.That(Convert.ToBoolean(loaded.Rows[0]["Active"]), Is.True);
        }

        [TestCase(".xlsx")]
        [TestCase(".xls")]
        public void SaveBookAndReadBackCorrectly(string extension)
        {
            XLBook source = new(
            [
                CreateSampleTable("First"),
                CreateSampleTable("Second"),
            ]);
            string filePath = Path.Combine(_tempDir, $"book{extension}");

            ExcelWriter.SaveBook(filePath, source, true, true);
            XLBook loaded = ExcelReader.GetBookParamsConnected(filePath);

            Assert.That(loaded.Sheets.Count, Is.EqualTo(2));
            Assert.That(loaded.Sheets[0].Rows.Count, Is.EqualTo(2));
            Assert.That(loaded.Sheets[1].Rows.Count, Is.EqualTo(2));
            Assert.That(loaded.Sheets[0].Rows[0]["Name"], Is.EqualTo("Alice"));
            Assert.That(loaded.Sheets[1].Rows[0]["Name"], Is.EqualTo("Alice"));
        }

        [Test]
        public void SaveBookParamsConnectedWritesAllSheets()
        {
            XLBook source = new(
            [
                CreateSampleTable("A"),
                CreateSampleTable("B"),
            ]);
            string filePath = Path.Combine(_tempDir, "connected.xlsx");

            ExcelWriter.SaveBookParamsConnected(filePath, source, writeColumnNames: true);
            XLBook loaded = ExcelReader.GetBookParamsConnected(filePath);

            Assert.That(loaded.Sheets.Count, Is.EqualTo(2));
            Assert.That(loaded.Sheets[1].Rows[1]["Name"], Is.EqualTo("Bob"));
        }

        private static DataTable CreateSampleTable(string tableName)
        {
            DataTable table = new(tableName);
            table.Columns.Add("Name", typeof(string));
            table.Columns.Add("Age", typeof(int));
            table.Columns.Add("Score", typeof(double));
            table.Columns.Add("Active", typeof(bool));

            table.Rows.Add("Alice", 30, 12.5d, true);
            table.Rows.Add("Bob", 25, 9.75d, false);
            return table;
        }
    }
}
