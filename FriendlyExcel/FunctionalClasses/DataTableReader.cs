// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel.FunctionalClasses
{
    internal class DataTableReader
    {
        public static IWorkbook Read(string filePath)
        {
            FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read);
                IWorkbook workbook;

                // Проверяем расширение файла и выбираем нужный класс
                if (filePath.EndsWith(".xls"))
                {
                    workbook = new HSSFWorkbook(fs); // Для старых файлов Excel (.xls)
                }
                else
                {
                    workbook = new XSSFWorkbook(fs); // Для новых файлов Excel (.xlsx)
                }
            return workbook;
        }
    }
}
