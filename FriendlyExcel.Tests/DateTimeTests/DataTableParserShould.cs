using Dumpify;
using FriendlyExcel.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel.Tests.DateTimeTests
{
    internal class DataTableParserShould
    {
        private static string TEST_DATA_PATH = ".\\test-data-time";
        private static string[] _test_paths = [];
        private IEnumerable<MethodInfo> methodInfos;
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
            Type type = assembly.GetTypes().Single(x => x.Name == "ExcelReader");
            Assert.IsNotNull(type);
            TypeInfo typeInfo = type.GetTypeInfo();
            Assert.IsNotNull(typeInfo);
            methodInfos = typeInfo.DeclaredMethods;
            methodInfos.Dump();
        }

        [Test]
        public void ReadDateTimeCorrectly([ValueSource(nameof(_test_paths))] string filePath, [Values(0)] int sheetIndex, [Values(true)] bool useFirstRowAsColumnNames)
        {
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "Get")!;
            Assert.IsNotNull(methodInfo);
            DataTable? table = (DataTable)methodInfo.Invoke(null, [filePath, sheetIndex, useFirstRowAsColumnNames])!;
            bool bookIsNotNull = table is not null;
            Assert.That(bookIsNotNull);
            table.Dump();
        }

    }


}
