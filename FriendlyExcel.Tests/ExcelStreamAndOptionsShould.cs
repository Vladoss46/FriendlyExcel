using FriendlyExcel.Models;
using System.Data;

namespace FriendlyExcel.Tests
{
    internal class ExcelStreamAndOptionsShould
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

        [TestCase(ExcelFormat.Xlsx)]
        [TestCase(ExcelFormat.Xls)]
        public void SaveAndGetViaStreamRoundTrip(ExcelFormat format)
        {
            DataTable source = CreateSampleTable("People");

            using MemoryStream stream = new();
            ExcelWriter.Save(stream, source, format);
            stream.Position = 0;

            DataTable loaded = ExcelReader.Get(stream, format);

            Assert.That(loaded.Rows.Count, Is.EqualTo(2));
            Assert.That(loaded.Rows[0]["Name"], Is.EqualTo("Alice"));
            Assert.That(loaded.TableName, Is.EqualTo("People"));
        }

        [Test]
        public void GetBySheetNameFromFile()
        {
            string filePath = Path.Combine(_tempDir, "named.xlsx");
            XLBook book = new(
            [
                CreateSampleTable("Orders"),
                CreateSampleTable("Customers"),
            ]);
            ExcelWriter.SaveBook(filePath, book);

            DataTable orders = ExcelReader.Get(filePath, "Orders");
            DataTable customers = ExcelReader.Get(filePath, sheetName: "customers"); // case-insensitive

            Assert.That(orders.TableName, Is.EqualTo("Orders"));
            Assert.That(customers.TableName, Is.EqualTo("Customers"));
            Assert.That(orders.Rows[0]["Name"], Is.EqualTo("Alice"));
        }

        [Test]
        public void GetWithReadOptionsSheetNameAndColumnTypes()
        {
            string filePath = Path.Combine(_tempDir, "options.xlsx");
            ExcelWriter.Save(filePath, CreateSampleTable("People"));

            DataTable table = ExcelReader.Get(filePath, new ExcelReadOptions
            {
                SheetName = "People",
                UseFirstRowAsColumnNames = true,
                ColumnTypes = [typeof(string), typeof(int), typeof(double), typeof(bool)],
            });

            Assert.That(table.Columns[1].DataType, Is.EqualTo(typeof(int)));
            Assert.That(Convert.ToInt32(table.Rows[0]["Age"]), Is.EqualTo(30));
        }

        [Test]
        public void GetBookFromStreamWithOptions()
        {
            XLBook source = new(
            [
                CreateSampleTable("A"),
                CreateSampleTable("B"),
            ]);

            using MemoryStream stream = new();
            ExcelWriter.SaveBook(stream, source, new ExcelWriteOptions
            {
                Format = ExcelFormat.Xlsx,
                WriteColumnNames = true,
            });
            stream.Position = 0;

            XLBook loaded = ExcelReader.GetBook(stream, new ExcelReadBookOptions
            {
                Format = ExcelFormat.Xlsx,
                UseFirstRowAsColumnNames = true,
            });

            Assert.That(loaded.Sheets.Count, Is.EqualTo(2));
            Assert.That(loaded.Sheets[0].TableName, Is.EqualTo("A"));
            Assert.That(loaded.Sheets[1].TableName, Is.EqualTo("B"));
        }

        [Test]
        public void GetFromStreamInfersFormatFromFileName()
        {
            DataTable source = CreateSampleTable("Upload");
            using MemoryStream stream = new();
            ExcelWriter.Save(stream, source, ExcelFormat.Xlsx);
            stream.Position = 0;

            DataTable loaded = ExcelReader.Get(stream, new ExcelReadOptions
            {
                FileName = "report.xlsx",
                SheetIndex = 0,
            });

            Assert.That(loaded.Rows.Count, Is.EqualTo(2));
        }

        [Test]
        public void GetUnknownSheetNameThrows()
        {
            string filePath = Path.Combine(_tempDir, "one.xlsx");
            ExcelWriter.Save(filePath, CreateSampleTable("Only"));

            Assert.Throws<ArgumentException>(() => ExcelReader.Get(filePath, "Missing"));
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
