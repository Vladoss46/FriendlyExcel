using NPOI.XSSF.UserModel;
using System.Data;
using System.Globalization;
using FlexibleNumberParser = FriendlyExcel.Internal.FlexibleNumberParser;

namespace FriendlyExcel.Tests
{
    internal class FlexibleNumberParsingShould
    {
        [TestCase("12,5", 12.5)]
        [TestCase("12.5", 12.5)]
        [TestCase("0,25", 0.25)]
        [TestCase("1000", 1000)]
        [TestCase("1.234,56", 1234.56)] // ru-style thousands + decimal
        [TestCase("1,234.56", 1234.56)] // en-style thousands + decimal
        public void ParsesDecimalCommaAndPoint(string text, double expected)
        {
            Assert.That(FlexibleNumberParser.TryParseDouble(text, out double value), Is.True);
            Assert.That(value, Is.EqualTo(expected).Within(0.0001));
        }

        [Test]
        public void DoesNotTreatRussianCommaTextAsOneHundredTwentyFive()
        {
            // InvariantCulture would parse "12,5" as 125 (comma = thousands separator).
            Assert.That(FlexibleNumberParser.TryParseDouble("12,5", out double value), Is.True);
            Assert.That(value, Is.EqualTo(12.5).Within(0.0001));
            Assert.That(value, Is.Not.EqualTo(125));
        }

        [Test]
        public void ReadsRussianTextNumbersFromExcelCellsAsDouble()
        {
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU");

            using var workbook = new XSSFWorkbook();
            var sheet = workbook.CreateSheet("Числа");
            var header = sheet.CreateRow(0);
            header.CreateCell(0).SetCellValue("Значение");
            header.CreateCell(1).SetCellValue("Сумма");

            var row = sheet.CreateRow(1);
            row.CreateCell(0).SetCellValue("12,5");   // text, Russian decimal comma
            row.CreateCell(1).SetCellValue(99.75);    // real numeric cell

            DataTable table = FriendlyExcel.FunctionalClasses.DataTableParser.Parse(sheet, useFirstRowAsColumnNames: true);

            Assert.That(table.Columns[0].DataType, Is.EqualTo(typeof(double)));
            Assert.That(table.Columns[1].DataType, Is.EqualTo(typeof(double)));
            Assert.That(Convert.ToDouble(table.Rows[0][0]), Is.EqualTo(12.5).Within(0.0001));
            Assert.That(Convert.ToDouble(table.Rows[0][1]), Is.EqualTo(99.75).Within(0.0001));
        }

        [Test]
        public void RoundTripTypedDoublesIndependentOfRussianUiSeparator()
        {
            // Excel stores typed numbers as binary doubles; UI may show "12,5" in ru-RU.
            CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("ru-RU");

            DataTable source = new("Лист1");
            source.Columns.Add("Цена", typeof(double));
            source.Rows.Add(12.5d);

            using MemoryStream stream = new();
            ExcelWriter.Save(stream, source, ExcelFormat.Xlsx);
            stream.Position = 0;

            DataTable loaded = ExcelReader.Get(stream, ExcelFormat.Xlsx);
            Assert.That(Convert.ToDouble(loaded.Rows[0]["Цена"]), Is.EqualTo(12.5).Within(0.0001));
        }
    }
}
