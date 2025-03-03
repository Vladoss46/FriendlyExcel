using FriendlyExcel.Exceptions;
using FriendlyExcel.Models;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;

//using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel
{
    internal class ExcelReader
    {
        public static DataTable Get(string filePath, int sheetIndex = 0, bool useFirstRowAsColumnNames = true)
        {
            using IWorkbook workbook = FunctionalClasses.DataTableReader.Read(filePath);
            var sheet = workbook.GetSheetAt(sheetIndex);
            DataTable table = FunctionalClasses.DataTableParser.Parse(sheet, useFirstRowAsColumnNames);
            workbook.Dispose();
            return table;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="useFirstRowAsColumnNames">a list for bool values of every table</param>
        /// <returns></returns>
        public static XLBook GetBook(string filePath, params bool[] useFirstRowAsColumnNames)
        {
            using IWorkbook workbook = FunctionalClasses.DataTableReader.Read(filePath);
            XLBook book = new();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                DataTable table;
                try
                {
                    table = FunctionalClasses.DataTableParser.Parse(sheet, useFirstRowAsColumnNames[i]);
                }
                catch (EmptySheetException)
                {
                    table = new DataTable();
                }
                catch (Exception)
                {
                    workbook.Dispose();
                    throw;
                }
                book.Sheets.Add(table);
            }
            workbook.Dispose();
            return book;
        }

        public static XLBook GetBookParamsConnected(string filePath, bool useFirstRowAsColumnNames)
        {
            using IWorkbook workbook = FunctionalClasses.DataTableReader.Read(filePath);
            XLBook book = new();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                var sheet = workbook.GetSheetAt(i);
                DataTable table;
                try
                {
                    table = FunctionalClasses.DataTableParser.Parse(sheet, useFirstRowAsColumnNames);
                }
                catch (EmptySheetException)
                {
                    table = new DataTable();
                }
                catch (Exception)
                {
                    workbook.Dispose();
                    throw;
                }
                book.Sheets.Add(table);
            }
            workbook.Dispose();
            return book;
        }
    }
}
