using FriendlyExcel.Models;
using System.Data;
using System.Globalization;

namespace FriendlyExcel.Tests
{
    /// <summary>
    /// Ensures Cyrillic text and non-invariant OS locales (ru-RU) do not break read/write.
    /// </summary>
    internal class CultureAndLocalizationShould
    {
        private CultureInfo _previousCulture = null!;
        private CultureInfo _previousUiCulture = null!;
        private string _tempDir = null!;

        [SetUp]
        public void Setup()
        {
            _previousCulture = CultureInfo.CurrentCulture;
            _previousUiCulture = CultureInfo.CurrentUICulture;
            _tempDir = Path.Combine(Path.GetTempPath(), "FriendlyExcelTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
        }

        [TearDown]
        public void TearDown()
        {
            CultureInfo.CurrentCulture = _previousCulture;
            CultureInfo.CurrentUICulture = _previousUiCulture;

            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, recursive: true);
        }

        [TestCase("ru-RU")]
        [TestCase("en-US")]
        [TestCase("de-DE")]
        public void RoundTripPreservesRussianTextUnderCulture(string cultureName)
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo(cultureName);
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo(cultureName);

            DataTable source = CreateRussianTable();
            string path = Path.Combine(_tempDir, $"ru-{cultureName}.xlsx");

            ExcelWriter.Save(path, source);
            DataTable loaded = ExcelReader.Get(path, "Сотрудники");

            Assert.That(loaded.TableName, Is.EqualTo("Сотрудники"));
            Assert.That(loaded.Columns[0].ColumnName, Is.EqualTo("Имя"));
            Assert.That(loaded.Columns[1].ColumnName, Is.EqualTo("Город"));
            Assert.That(loaded.Columns[2].ColumnName, Is.EqualTo("Зарплата"));
            Assert.That(loaded.Columns[3].ColumnName, Is.EqualTo("Активен"));

            Assert.That(loaded.Rows[0]["Имя"], Is.EqualTo("Иван Петров"));
            Assert.That(loaded.Rows[0]["Город"], Is.EqualTo("Москва"));
            Assert.That(Convert.ToDouble(loaded.Rows[0]["Зарплата"]), Is.EqualTo(125000.5).Within(0.001));
            Assert.That(Convert.ToBoolean(loaded.Rows[0]["Активен"]), Is.True);

            Assert.That(loaded.Rows[1]["Имя"], Is.EqualTo("Марія Коваленко"));
            Assert.That(loaded.Rows[1]["Город"], Is.EqualTo("Київ"));
        }

        [Test]
        public void StreamRoundTripUnderRussianCultureKeepsNumericTypesStable()
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU");
            CultureInfo.CurrentUICulture = CultureInfo.GetCultureInfo("ru-RU");

            DataTable source = CreateRussianTable();
            using MemoryStream stream = new();
            ExcelWriter.Save(stream, source, ExcelFormat.Xlsx);
            stream.Position = 0;

            DataTable loaded = ExcelReader.Get(stream, new ExcelReadOptions
            {
                Format = ExcelFormat.Xlsx,
                SheetName = "Сотрудники",
            });

            Assert.That(loaded.Columns["Зарплата"]!.DataType, Is.EqualTo(typeof(double)));
            Assert.That(Convert.ToDouble(loaded.Rows[0]["Зарплата"], CultureInfo.InvariantCulture), Is.EqualTo(125000.5).Within(0.001));
            // Must not become a locale-formatted string like "125000,5"
            Assert.That(loaded.Rows[0]["Зарплата"], Is.Not.TypeOf<string>());
        }

        [Test]
        public void ExistingRussianHeadersInTestDataAreReadCorrectly()
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU");

            string path = Path.Combine(AppContext.BaseDirectory, "test-data", "test.xlsx");
            DataTable table = ExcelReader.Get(path);

            Assert.That(table.Columns[0].ColumnName, Is.EqualTo("Строки"));
            Assert.That(table.Columns[1].ColumnName, Is.EqualTo("Числа"));
            Assert.That(table.Columns[2].ColumnName, Is.EqualTo("Буквы"));
            Assert.That(table.Columns[3].ColumnName, Is.EqualTo("Дробные"));
            Assert.That(table.Columns[4].ColumnName, Is.EqualTo("Логические"));
            Assert.That(table.Rows.Count, Is.GreaterThan(0));
        }

        [Test]
        public void SheetNameLookupIsCaseInsensitiveForCyrillic()
        {
            DataTable source = CreateRussianTable();
            string path = Path.Combine(_tempDir, "case.xlsx");
            ExcelWriter.Save(path, source);

            DataTable loaded = ExcelReader.Get(path, "сотрудники");
            Assert.That(loaded.Rows[0]["Имя"], Is.EqualTo("Иван Петров"));
        }

        private static DataTable CreateRussianTable()
        {
            DataTable table = new("Сотрудники");
            table.Columns.Add("Имя", typeof(string));
            table.Columns.Add("Город", typeof(string));
            table.Columns.Add("Зарплата", typeof(double));
            table.Columns.Add("Активен", typeof(bool));

            table.Rows.Add("Иван Петров", "Москва", 125000.5d, true);
            table.Rows.Add("Марія Коваленко", "Київ", 98000.25d, false);
            return table;
        }
    }
}
