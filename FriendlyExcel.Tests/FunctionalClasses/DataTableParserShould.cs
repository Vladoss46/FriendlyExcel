// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using NPOI.SS.UserModel;
using System.Data;
using ExcelDataTableParser = FriendlyExcel.FunctionalClasses.DataTableParser;
using ExcelDataTableReader = FriendlyExcel.FunctionalClasses.DataTableReader;

namespace FriendlyExcel.Tests.FunctionalClasses
{
    internal class DataTableParserShould
    {
        private IWorkbook? _workbook;
        private static readonly string TestDataPath = Path.Combine(AppContext.BaseDirectory, "test-data");
        private static readonly string[] TestPaths = Directory.GetFiles(TestDataPath);

        private static readonly string[] TestColumnNames =
        [
            "Строки",
            "Числа",
            "Буквы",
            "Дробные",
            "Логические",
        ];

        private static readonly string[] TestBaseColumnNames =
        [
            "Column1",
            "Column2",
            "Column3",
            "Column4",
            "Column5",
        ];

        private static readonly Type[] TestColumnTypes =
        [
            typeof(string),
            typeof(int),
            typeof(string),
            typeof(double),
            typeof(string),
        ];

        private static readonly (CellType, Type)[] CellTypes =
        [
            (CellType.Boolean, typeof(bool)),
            (CellType.Numeric, typeof(double)),
            (CellType.String, typeof(string)),
            (CellType.Blank, typeof(string)),
            (CellType.Unknown, typeof(string)),
            (CellType.Error, typeof(string)),
            (CellType.Formula, typeof(string)),
        ];

        private static readonly (double Value, Type Type)[] DoubleTypes =
        [
            (123d, typeof(int)),
            (12.3d, typeof(double)),
            (-5d, typeof(int)),
        ];

        private static readonly (Type[] Current, Type[] Accumulated, Type[] Result)[] ColumnTypeArrays =
        [
            (
                [typeof(int), typeof(string), typeof(double)],
                [],
                [typeof(int), typeof(string), typeof(double)]
            ),
            (
                [typeof(string), typeof(string), typeof(string)],
                [typeof(int), typeof(string), typeof(double)],
                [typeof(string), typeof(string), typeof(string)]
            ),
            (
                [typeof(int), typeof(string), typeof(double)],
                [typeof(string), typeof(string), typeof(string)],
                [typeof(string), typeof(string), typeof(string)]
            ),
            (
                [typeof(double), typeof(int)],
                [typeof(int), typeof(double)],
                [typeof(double), typeof(double)]
            ),
        ];

        private static readonly (int SheetIndex, int StartRowIndex)[] SheetAndRowCombinations =
        [
            (0, 1),
            (1, 0),
        ];

        private static readonly (int CellIndex, Type CellType)[] CellSource =
        [
            (0, typeof(string)),
            (1, typeof(int)),
            (2, typeof(string)),
            (3, typeof(double)),
            (4, typeof(string)),
        ];

        [TearDown]
        public void TearDown()
        {
            _workbook?.Dispose();
            _workbook = null;
        }

        [Test]
        [TestCaseSource(nameof(TestPaths))]
        public void GetColumnNamesCorrectly(string filePath)
        {
            ISheet sheet = OpenSheet(filePath, 0);
            string[] columnNames = ExcelDataTableParser.GetColumnNames(sheet);
            Assert.That(columnNames, Is.EquivalentTo(TestColumnNames));
        }

        [Test]
        [TestCaseSource(nameof(TestPaths))]
        public void GetBaseColumnNamesCorrectly(string filePath)
        {
            ISheet sheet = OpenSheet(filePath, 0);
            string[] columnNames = ExcelDataTableParser.GetBaseColumnNames(sheet);
            Assert.That(columnNames, Is.EquivalentTo(TestBaseColumnNames));
        }

        [Test]
        [TestCaseSource(nameof(CellTypes))]
        public void GetTypeOfCellCorrectly((CellType CellType, Type Expected) type)
        {
            Type netType = ExcelDataTableParser.GetTypeOfCell(type.CellType);
            Assert.That(netType, Is.EqualTo(type.Expected));
        }

        [Test]
        [TestCaseSource(nameof(DoubleTypes))]
        public void CheckDoubleTypeCorrectly((double Value, Type Type) type)
        {
            Type netType = ExcelDataTableParser.CheckDoubleType(type.Value);
            Assert.That(netType, Is.EqualTo(type.Type));
        }

        [Test]
        [TestCaseSource(nameof(ColumnTypeArrays))]
        public void CompareAndSwitchCorrectly((Type[] Current, Type[] Accumulated, Type[] Result) data)
        {
            Type[] resultTypes = ExcelDataTableParser.CompareAndSwitch(data.Current, data.Accumulated);
            Assert.That(resultTypes, Is.EqualTo(data.Result));
        }

        [Test, Combinatorial]
        public void GetColumnTypesCorrectly(
            [ValueSource(nameof(TestPaths))] string filePath,
            [ValueSource(nameof(SheetAndRowCombinations))] (int SheetIndex, int StartRowIndex) combo)
        {
            ISheet sheet = OpenSheet(filePath, combo.SheetIndex);
            Type[] columnTypes = ExcelDataTableParser.GetColumnTypes(sheet, combo.StartRowIndex);
            Assert.That(columnTypes, Is.EquivalentTo(TestColumnTypes));
        }

        [Test]
        public void CreateDataTableCorrectly()
        {
            DataTable table = ExcelDataTableParser.CreateDataTable(TestColumnNames, TestColumnTypes);
            Assert.That(table.Columns.Count, Is.EqualTo(TestColumnNames.Length));
            Assert.That(table.Columns[1].DataType, Is.EqualTo(typeof(int)));
        }

        [Test]
        public void GetCellValueCorrectly(
            [ValueSource(nameof(TestPaths))] string filePath,
            [ValueSource(nameof(CellSource))] (int CellIndex, Type CellType) cellSource)
        {
            ISheet sheet = OpenSheet(filePath, 1);
            ICell cell = sheet.GetRow(0).GetCell(cellSource.CellIndex);
            object value = ExcelDataTableParser.GetCellValue(cell, cellSource.CellType);
            Assert.That(value, Is.Not.Null);
        }

        [Test]
        public void ParseRowCorrectly([ValueSource(nameof(TestPaths))] string filePath)
        {
            ISheet sheet = OpenSheet(filePath, 1);
            IRow row = sheet.GetRow(0);
            DataTable table = ExcelDataTableParser.CreateDataTable(TestColumnNames, TestColumnTypes);
            DataRow tableRow = table.NewRow();

            DataRow result = ExcelDataTableParser.ParseRow(row, tableRow, TestColumnTypes);
            Assert.That(result, Is.Not.Null);
            Assert.That(result[0], Is.Not.EqualTo(DBNull.Value));
        }

        [Test]
        public void FillTableCorrectly(
            [ValueSource(nameof(TestPaths))] string filePath,
            [ValueSource(nameof(SheetAndRowCombinations))] (int SheetIndex, int StartRowIndex) combo)
        {
            ISheet sheet = OpenSheet(filePath, combo.SheetIndex);
            DataTable table = ExcelDataTableParser.CreateDataTable(TestColumnNames, TestColumnTypes);

            table = ExcelDataTableParser.FillTable(table, sheet, TestColumnTypes, combo.StartRowIndex);

            Assert.That(table.Rows.Count, Is.EqualTo(sheet.PhysicalNumberOfRows - combo.StartRowIndex));
        }

        [Test]
        public void ParseCorrectly(
            [ValueSource(nameof(TestPaths))] string filePath,
            [ValueSource(nameof(SheetAndRowCombinations))] (int SheetIndex, int StartRowIndex) combo)
        {
            ISheet sheet = OpenSheet(filePath, combo.SheetIndex);
            bool useFirstRowAsColumnNames = combo.StartRowIndex == 1;

            DataTable table = ExcelDataTableParser.Parse(sheet, useFirstRowAsColumnNames);

            Assert.That(table, Is.Not.Null);
            Assert.That(table.Rows.Count, Is.EqualTo(sheet.PhysicalNumberOfRows - combo.StartRowIndex));
            Assert.That(table.Columns[1].DataType, Is.EqualTo(typeof(int)));
        }

        private ISheet OpenSheet(string filePath, int sheetIndex)
        {
            _workbook?.Dispose();
            _workbook = ExcelDataTableReader.Read(filePath);
            return _workbook.GetSheetAt(sheetIndex);
        }
    }
}
