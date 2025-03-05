// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using Dumpify;
using FriendlyExcel.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel.Tests
{
    internal class ExcelReaderShould//TODO: проверь работу его методов на test-data-sheets
    {
        private static string TEST_DATA_PATH = ".\\test-data-sheets";
        private static string[] _test_paths = [];
        private IEnumerable<MethodInfo> methodInfos;
        static ExcelReaderShould()
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
        public void GetCorrectly([ValueSource(nameof(_test_paths))] string filePath, [Values(0, 1)] int sheetIndex, [Values(true)] bool useFirstRowAsColumnNames)
        {
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "Get")!;
            Assert.IsNotNull(methodInfo);
            DataTable? table = (DataTable)methodInfo.Invoke(null, [filePath, sheetIndex, useFirstRowAsColumnNames])!;
            bool bookIsNotNull = table is not null;
            Assert.That(bookIsNotNull);
            table.Dump();
        }
        private static IEnumerable<TestCaseData> GetTestData1()
        {
            yield return new TestCaseData(_test_paths.Single(x=>x.EndsWith(".xls")),true,false);
            yield return new TestCaseData(_test_paths.Single(x=>x.EndsWith(".xlsx")),true,true);
        }
        [Test]
        [TestCaseSource(nameof(GetTestData1))]
        public void GetBookCorrectly(string filePaths, bool useFirstRowAsColumnNames1, bool useFirstRowAsColumnNames2)
        {
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "GetBook")!;
            Assert.IsNotNull(methodInfo);
            XLBook? book = (XLBook)methodInfo.Invoke(null, [filePaths, new List<bool> { useFirstRowAsColumnNames1, useFirstRowAsColumnNames2 }.ToArray()])!;
            bool bookIsNotNull = book is not null;
            Assert.That(bookIsNotNull);
            book.Dump();
        }

        [Test]
        public void GetBookParamsConnectedCorrectly([ValueSource(nameof(_test_paths))] string filePath, [Values(true)] bool useFirstRowAsColumnNames)
        {
            MethodInfo methodInfo = methodInfos.Single(x => x.Name == "GetBookParamsConnected")!;
            Assert.IsNotNull(methodInfo);
            XLBook? book = (XLBook)methodInfo.Invoke(null, [filePath, useFirstRowAsColumnNames])!;
            bool bookIsNotNull = book is not null;
            Assert.That(bookIsNotNull);
            book.Dump();
        }

    }
}
