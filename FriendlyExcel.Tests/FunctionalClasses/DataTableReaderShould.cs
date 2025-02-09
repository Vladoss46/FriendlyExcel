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

namespace FriendlyExcel.Tests.FunctionalClasses
{
    internal class DataTableReaderShould
    {
        private static string TEST_DATA_PATH= ".\\test-data";
        private static string[] _test_paths = [];
        private IEnumerable<MethodInfo> methodInfos;
        static DataTableReaderShould()
        {
            _test_paths = [.. Directory.GetFiles(TEST_DATA_PATH)];
        }

        [SetUp]
        public void Setup()
        {
            Assembly assembly = Assembly.GetAssembly(typeof(XLBook))!;
            Assert.IsNotNull(assembly);
            assembly.GetTypes().Dump();
            Type type = assembly.GetTypes().Single(x => x.Name == "DataTableReader");
            Assert.IsNotNull(type);
            TypeInfo typeInfo = type.GetTypeInfo();
            Assert.IsNotNull(typeInfo);
            methodInfos = typeInfo.DeclaredMethods;
            methodInfos.Dump();
        }

        [Test]
        [TestCaseSource(nameof(_test_paths))]
        public void ReadCorrectly(string filePath)
        {
            MethodInfo methodInfo = methodInfos.Single(x=>x.Name == "Read")!;
            Assert.IsNotNull(methodInfo);
            IWorkbook book = (IWorkbook)methodInfo.Invoke(null, [filePath])!;
            bool bookIsNotNull = book is not null;
            book?.Dispose();
            Assert.That(bookIsNotNull);
        }
    }
}
