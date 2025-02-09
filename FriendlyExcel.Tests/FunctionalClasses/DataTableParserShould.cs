// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.


using Dumpify;
using FriendlyExcel.Models;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NUnit.Framework.Constraints;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel.Tests.FunctionalClasses
{
    internal class DataTableParserShould
    {
        private IWorkbook workbook;
        private ISheet sheet;
        private static string TEST_DATA_PATH = ".\\test-data";
        private static string[] _test_paths = [];
        private IEnumerable<MethodInfo> methodInfos;
        private MethodInfo readMethodInfo;
        static DataTableParserShould()
        {
            _test_paths = [.. Directory.GetFiles(TEST_DATA_PATH)];
        }

        [SetUp]
        public void Setup()
        {
            Assembly assembly = Assembly.GetAssembly(typeof(XLBook))!;
            Assert.IsNotNull(assembly);
            assembly.GetTypes().Dump();
            Type type = assembly.GetTypes().Single(x => x.Name == "DataTableParser");
            Assert.IsNotNull(type);
            TypeInfo typeInfo = type.GetTypeInfo();
            Assert.IsNotNull(typeInfo);
            methodInfos = typeInfo.DeclaredMethods;
            methodInfos.Dump();

            readMethodInfo = assembly.GetTypes().Single(x => x.Name == "DataTableReader").GetTypeInfo().DeclaredMethods.Single(x => x.Name == "Read")!;
        }
        private static string[] TEST_COLUMN_NAMES = [
            "Строки",
            "Числа",
            "Буквы",
            "Дробные",
            "Логические"];
        [Test]
        [TestCaseSource(nameof(_test_paths))]
        public void GetColumnNamesCorrectly(string filePath)
        {
            workbook = (IWorkbook)readMethodInfo.Invoke(null, [filePath])!;
            sheet = workbook.GetSheetAt(0);
            Assert.IsNotNull(sheet);
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "GetColumnNames")!;
            Assert.IsNotNull(methodInfo);
            var columnNames = (string[])methodInfo.Invoke(null, [sheet])!;
            columnNames.Dump();
            Assert.That(columnNames, Is.EquivalentTo(TEST_COLUMN_NAMES));

        }
        private static string[] TEST_BASE_COLUMN_NAMES = [
            "Column1",
            "Column2",
            "Column3",
            "Column4",
            "Column5"];
        [Test]
        [TestCaseSource(nameof(_test_paths))]
        public void GetBaseColumnNamesCorrectly(string filePath)
        {
            workbook = (IWorkbook)readMethodInfo.Invoke(null, [filePath])!;
            sheet = workbook.GetSheetAt(0);
            Assert.IsNotNull(sheet);
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "GetBaseColumnNames")!;
            Assert.IsNotNull(methodInfo);
            var columnNames = (string[])methodInfo.Invoke(null, [sheet])!;
            columnNames.Dump();
            Assert.That(columnNames, Is.EquivalentTo(TEST_BASE_COLUMN_NAMES));

        }
        private static (CellType, Type)[] TYPES = [
            (CellType.Boolean,typeof(Boolean)),
            (CellType.Numeric,typeof(double)),
            (CellType.String,typeof(string)),
            (CellType.Blank,typeof(string)),
            (CellType.Unknown,typeof(string)),
            (CellType.Error,typeof(string)),
            (CellType.Formula,typeof(string)),
            ];
        [Test]
        [TestCaseSource(nameof(TYPES))]
        public void GetTypeOfCellCorrectly((CellType, Type) type)
        {
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "<GetColumnTypes>g__GetTypeOfCell|3_0")!;
            Assert.IsNotNull(methodInfo);
            var net_type = (Type)methodInfo.Invoke(null, [type.Item1])!;
            net_type.Dump();
            Assert.That(net_type, Is.EqualTo(type.Item2));

        }
        private static (string, Type)[] DOUBLE_TYPES = [
            ("123",typeof(int)),
            ("12,3",typeof(double)),
            ("12.3",typeof(double)),
            ];
        [Test]
        [TestCaseSource(nameof(DOUBLE_TYPES))]
        public void CheckDoubleTypeCorrectly((string value, Type type) type)
        {
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "<GetColumnTypes>g__CheckDoubleType|3_1")!;
            Assert.IsNotNull(methodInfo);
            var net_type = (Type)methodInfo.Invoke(null, [type.value])!;
            net_type.Dump();
            Assert.That(net_type, Is.EqualTo(type.type));

        }
        private static (Type[], Type[], Type[])[] COLUMN_TYPE_ARRAYS = [
           ( [
               typeof(int),
               typeof(string),
               typeof(double)
               ]
            ,
            [],
            [
               typeof(int),
               typeof(string),
               typeof(double)
            ]),
            ///////////////
            (  [
               typeof(string),
               typeof(string),
               typeof(string)
               ]
            ,
               [
               typeof(int),
               typeof(string),
               typeof(double)
               ]
            ,
               [
               typeof(string),
               typeof(string),
               typeof(string)
            ]),
            ///////////////
            ( [
               typeof(int),
               typeof(string),
               typeof(double)
               ]
            ,
               [
               typeof(string),
               typeof(string),
               typeof(string)
               ]
            ,
               [
               typeof(string),
               typeof(string),
               typeof(string)
            ]),
            ];
        [Test]
        [TestCaseSource(nameof(COLUMN_TYPE_ARRAYS))]
        public void CompareAndSwitchCorrectly((Type[] currentColumnTypes, Type[] columnTypes, Type[] resultColumnTypes) _)
        {
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "<GetColumnTypes>g__CompareAndSwitch|3_2")!;
            Assert.IsNotNull(methodInfo);
            var resultTypes = (Type[])methodInfo.Invoke(null, [_.currentColumnTypes, _.columnTypes])!;
            resultTypes.Dump();
            Assert.That(resultTypes, Is.EqualTo(_.resultColumnTypes));
        }
        private static (int, int)[] TEST_COLUMN_TYPES_SHEET_AND_ROW_COMBINATIONS = [
            (0,1),
            (1,0)];
        private static Type[] TEST_COLUMN_TYPES = [
           typeof(string),
           typeof(int),
           typeof(string),
           typeof(double),
           typeof(string),
            ];
        [Test, Combinatorial]
        public void GetColumnTypesCorrectly([ValueSource(nameof(_test_paths))]string filePath, [ValueSource(nameof(TEST_COLUMN_TYPES_SHEET_AND_ROW_COMBINATIONS))](int sheetIndex, int startRowIndex) _)
        {
            workbook = (IWorkbook)readMethodInfo.Invoke(null, [filePath])!;
            sheet = workbook.GetSheetAt(_.sheetIndex);
            Assert.IsNotNull(sheet);
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "GetColumnTypes")!;
            Assert.IsNotNull(methodInfo);
            var columnTypes = (Type[])methodInfo.Invoke(null, [sheet,_.startRowIndex])!;
            columnTypes.Dump();
            Assert.That(columnTypes, Is.EquivalentTo(TEST_COLUMN_TYPES));

        }
        private static (string[] names, Type[] types)[] TEST_COLUMN_SOURCE =
            [(TEST_COLUMN_NAMES,TEST_COLUMN_TYPES)];
        [Test]
        public void CreateDataTableCorrectly( [ValueSource(nameof(TEST_COLUMN_SOURCE))](string [] columnNames, Type[] columnTypes)_)
        {
            //workbook = (IWorkbook)readMethodInfo.Invoke(null, [filePath])!;
            //sheet = workbook.GetSheetAt(_.sheetIndex);
            //Assert.IsNotNull(sheet);
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "CreateDataTable")!;
            Assert.IsNotNull(methodInfo);
            var table = (DataTable)methodInfo.Invoke(null, [_.columnNames, _.columnTypes])!;
            table.Dump();
            Assert.That(table, Is.Not.Null);

        }
        [TearDown]
        public void TearDown()
        {
            workbook?.Dispose();
        }
    }
}
