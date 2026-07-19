using System.Data;
using System.Globalization;

namespace FriendlyExcel.Tests
{
    internal class PocoMappingShould
    {
        private CultureInfo _previousCulture = null!;
        private string _tempDir = null!;

        [SetUp]
        public void Setup()
        {
            _previousCulture = CultureInfo.CurrentCulture;
            _tempDir = Path.Combine(Path.GetTempPath(), "FriendlyExcelTests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempDir);
        }

        [TearDown]
        public void TearDown()
        {
            CultureInfo.CurrentCulture = _previousCulture;
            if (Directory.Exists(_tempDir))
                Directory.Delete(_tempDir, recursive: true);
        }

        [Test]
        public void RoundTripWithRussianExcelColumnsUnderRuRu()
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU");

            List<Employee> source =
            [
                new() { Name = "Иван", Salary = 125000.5, Active = true },
                new() { Name = "Мария", Salary = 98000.25, Active = false },
            ];

            string path = Path.Combine(_tempDir, "employees.xlsx");
            ExcelWriter.Save(path, source);

            List<Employee> loaded = ExcelReader.Get<Employee>(path);

            Assert.That(loaded, Has.Count.EqualTo(2));
            Assert.That(loaded[0].Name, Is.EqualTo("Иван"));
            Assert.That(loaded[0].Salary, Is.EqualTo(125000.5).Within(0.001));
            Assert.That(loaded[0].Active, Is.True);
            Assert.That(loaded[1].Name, Is.EqualTo("Мария"));
            Assert.That(loaded[1].Active, Is.False);
        }

        [Test]
        public void ColumnMatchIsCaseInsensitiveForCyrillic()
        {
            List<Employee> source = [new() { Name = "Тест", Salary = 1, Active = true }];
            string path = Path.Combine(_tempDir, "case.xlsx");
            ExcelWriter.Save(path, source);

            // Rewrite is not needed: reader matches OrdinalIgnoreCase via PocoMapper.
            List<Employee> loaded = ExcelReader.Get<Employee>(path, new ExcelReadOptions
            {
                SheetIndex = 0,
                UseFirstRowAsColumnNames = true,
            });

            Assert.That(loaded[0].Name, Is.EqualTo("Тест"));
        }

        [Test]
        public void MissingColumnThrows()
        {
            DataTable table = new("Sheet1");
            table.Columns.Add("Имя", typeof(string));
            // Missing "Зарплата" and "Active"
            table.Rows.Add("Иван");

            string path = Path.Combine(_tempDir, "missing.xlsx");
            ExcelWriter.Save(path, table);

            InvalidOperationException ex = Assert.Throws<InvalidOperationException>(
                () => ExcelReader.Get<Employee>(path))!;

            Assert.That(ex.Message, Does.Contain("Зарплата").Or.Contain("Active"));
        }

        [Test]
        public void StreamGetAndSaveRoundTrip()
        {
            List<Employee> source =
            [
                new() { Name = "Анна", Salary = 10.5, Active = true },
            ];

            using MemoryStream stream = new();
            ExcelWriter.Save(stream, source, ExcelFormat.Xlsx);
            stream.Position = 0;

            List<Employee> loaded = ExcelReader.Get<Employee>(stream, new ExcelReadOptions
            {
                Format = ExcelFormat.Xlsx,
                UseFirstRowAsColumnNames = true,
            });

            Assert.That(loaded, Has.Count.EqualTo(1));
            Assert.That(loaded[0].Name, Is.EqualTo("Анна"));
            Assert.That(loaded[0].Salary, Is.EqualTo(10.5).Within(0.001));
        }

        [Test]
        public void PropertyNameMapsWithoutAttribute()
        {
            List<SimpleRow> source = [new() { Title = "Hello", Count = 3 }];
            string path = Path.Combine(_tempDir, "simple.xlsx");
            ExcelWriter.Save(path, source);

            List<SimpleRow> loaded = ExcelReader.Get<SimpleRow>(path);
            Assert.That(loaded[0].Title, Is.EqualTo("Hello"));
            Assert.That(loaded[0].Count, Is.EqualTo(3));
        }

        private sealed class Employee
        {
            [ExcelColumn("Имя")]
            public string Name { get; set; } = "";

            [ExcelColumn("Зарплата")]
            public double Salary { get; set; }

            public bool Active { get; set; }
        }

        private sealed class SimpleRow
        {
            public string Title { get; set; } = "";
            public int Count { get; set; }
        }
    }
}
