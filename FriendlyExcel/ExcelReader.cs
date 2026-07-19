// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using FriendlyExcel.Exceptions;
using FriendlyExcel.Internal;
using FriendlyExcel.Models;
using NPOI.SS.UserModel;
using System.Data;

namespace FriendlyExcel
{
    /// <summary>
    /// Reads Excel workbooks (.xls / .xlsx) into <see cref="DataTable"/> / <see cref="XLBook"/>.
    /// </summary>
    public static class ExcelReader
    {
        /// <summary>
        /// Reads a single sheet into a <see cref="DataTable"/>.
        /// </summary>
        public static DataTable Get(string filePath, int sheetIndex = 0, bool useFirstRowAsColumnNames = true)
        {
            return Get(filePath, new ExcelReadOptions
            {
                SheetIndex = sheetIndex,
                UseFirstRowAsColumnNames = useFirstRowAsColumnNames,
            });
        }

        /// <summary>
        /// Reads a single sheet by name into a <see cref="DataTable"/>.
        /// </summary>
        public static DataTable Get(string filePath, string sheetName, bool useFirstRowAsColumnNames = true)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
            return Get(filePath, new ExcelReadOptions
            {
                SheetName = sheetName,
                UseFirstRowAsColumnNames = useFirstRowAsColumnNames,
            });
        }

        /// <summary>
        /// Reads a single sheet using explicit column types.
        /// </summary>
        public static DataTable Get(string filePath, Type[] columnTypes, int sheetIndex = 0, bool useFirstRowAsColumnNames = true)
        {
            return Get(filePath, new ExcelReadOptions
            {
                SheetIndex = sheetIndex,
                UseFirstRowAsColumnNames = useFirstRowAsColumnNames,
                ColumnTypes = columnTypes,
            });
        }

        /// <summary>
        /// Reads a single sheet from a file using <see cref="ExcelReadOptions"/>.
        /// </summary>
        public static DataTable Get(string filePath, ExcelReadOptions options)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            ArgumentNullException.ThrowIfNull(options);

            using IWorkbook workbook = FunctionalClasses.DataTableReader.Read(filePath);
            return ParseSheet(workbook, options);
        }

        /// <summary>
        /// Reads a single sheet from a stream. <see cref="ExcelReadOptions.Format"/> or <see cref="ExcelReadOptions.FileName"/> is required.
        /// The stream is not disposed.
        /// </summary>
        public static DataTable Get(Stream stream, ExcelReadOptions options)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(options);

            ExcelFormat format = ExcelFormatResolver.Resolve(options.Format, options.FileName);
            using IWorkbook workbook = FunctionalClasses.DataTableReader.Read(stream, format);
            return ParseSheet(workbook, options);
        }

        /// <summary>
        /// Reads a single sheet from a stream with an explicit format.
        /// </summary>
        public static DataTable Get(Stream stream, ExcelFormat format, int sheetIndex = 0, bool useFirstRowAsColumnNames = true)
        {
            return Get(stream, new ExcelReadOptions
            {
                Format = format,
                SheetIndex = sheetIndex,
                UseFirstRowAsColumnNames = useFirstRowAsColumnNames,
            });
        }

        /// <summary>
        /// Reads a single sheet by name from a stream with an explicit format.
        /// </summary>
        public static DataTable Get(Stream stream, ExcelFormat format, string sheetName, bool useFirstRowAsColumnNames = true)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);
            return Get(stream, new ExcelReadOptions
            {
                Format = format,
                SheetName = sheetName,
                UseFirstRowAsColumnNames = useFirstRowAsColumnNames,
            });
        }

        /// <summary>
        /// Reads every sheet into an <see cref="XLBook"/>.
        /// </summary>
        /// <param name="filePath">Path to .xls or .xlsx file.</param>
        /// <param name="useFirstRowAsColumnNames">Per-sheet flags; missing values default to <c>true</c>.</param>
        public static XLBook GetBook(string filePath, params bool[] useFirstRowAsColumnNames)
        {
            return GetBook(filePath, new ExcelReadBookOptions
            {
                UseFirstRowAsColumnNames = true,
                UseFirstRowAsColumnNamesPerSheet = useFirstRowAsColumnNames.Length > 0 ? useFirstRowAsColumnNames : null,
            });
        }

        /// <summary>
        /// Reads every sheet with the same header option for all sheets.
        /// </summary>
        public static XLBook GetBookParamsConnected(string filePath, bool useFirstRowAsColumnNames = true)
        {
            return GetBook(filePath, new ExcelReadBookOptions
            {
                UseFirstRowAsColumnNames = useFirstRowAsColumnNames,
            });
        }

        /// <summary>
        /// Reads every sheet from a file using <see cref="ExcelReadBookOptions"/>.
        /// </summary>
        public static XLBook GetBook(string filePath, ExcelReadBookOptions options)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            ArgumentNullException.ThrowIfNull(options);

            using IWorkbook workbook = FunctionalClasses.DataTableReader.Read(filePath);
            return ParseBook(workbook, options);
        }

        /// <summary>
        /// Reads every sheet from a stream. <see cref="ExcelReadBookOptions.Format"/> or <see cref="ExcelReadBookOptions.FileName"/> is required.
        /// The stream is not disposed.
        /// </summary>
        public static XLBook GetBook(Stream stream, ExcelReadBookOptions options)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(options);

            ExcelFormat format = ExcelFormatResolver.Resolve(options.Format, options.FileName);
            using IWorkbook workbook = FunctionalClasses.DataTableReader.Read(stream, format);
            return ParseBook(workbook, options);
        }

        /// <summary>
        /// Reads every sheet from a stream with an explicit format.
        /// </summary>
        public static XLBook GetBook(Stream stream, ExcelFormat format, bool useFirstRowAsColumnNames = true)
        {
            return GetBook(stream, new ExcelReadBookOptions
            {
                Format = format,
                UseFirstRowAsColumnNames = useFirstRowAsColumnNames,
            });
        }

        private static DataTable ParseSheet(IWorkbook workbook, ExcelReadOptions options)
        {
            ISheet sheet = WorkbookSheetSelector.GetSheet(workbook, options);
            if (options.ColumnTypes is { Length: > 0 })
                return FunctionalClasses.DataTableParser.Parse(sheet, options.ColumnTypes, options.UseFirstRowAsColumnNames);

            return FunctionalClasses.DataTableParser.Parse(sheet, options.UseFirstRowAsColumnNames);
        }

        private static XLBook ParseBook(IWorkbook workbook, ExcelReadBookOptions options)
        {
            XLBook book = new();
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                ISheet sheet = workbook.GetSheetAt(i);
                bool useHeaders = ResolveHeaderFlag(options, i);

                try
                {
                    DataTable table = FunctionalClasses.DataTableParser.Parse(sheet, useHeaders);
                    book.Sheets.Add(table);
                }
                catch (EmptySheetException)
                {
                    book.Sheets.Add(new DataTable { TableName = sheet.SheetName });
                }
            }

            return book;
        }

        private static bool ResolveHeaderFlag(ExcelReadBookOptions options, int sheetIndex)
        {
            if (options.UseFirstRowAsColumnNamesPerSheet is { Length: > 0 } perSheet
                && sheetIndex < perSheet.Length)
            {
                return perSheet[sheetIndex];
            }

            return options.UseFirstRowAsColumnNames;
        }
    }
}
