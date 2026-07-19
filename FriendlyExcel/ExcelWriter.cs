// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using FriendlyExcel.Internal;
using FriendlyExcel.Models;
using NPOI.SS.UserModel;
using System.Data;

namespace FriendlyExcel
{
    /// <summary>
    /// Writes <see cref="DataTable"/> / <see cref="XLBook"/> into Excel files (.xls / .xlsx).
    /// </summary>
    public static class ExcelWriter
    {
        /// <summary>
        /// Saves a single <see cref="DataTable"/> as an Excel file.
        /// </summary>
        public static void Save(string filePath, DataTable table, bool writeColumnNames = true)
        {
            Save(filePath, table, new ExcelWriteOptions { WriteColumnNames = writeColumnNames });
        }

        /// <summary>
        /// Saves a single <see cref="DataTable"/> to a file using <see cref="ExcelWriteOptions"/>.
        /// </summary>
        public static void Save(string filePath, DataTable table, ExcelWriteOptions options)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            ArgumentNullException.ThrowIfNull(table);
            ArgumentNullException.ThrowIfNull(options);

            ExcelFormat format = ExcelFormatResolver.Resolve(options.Format, filePath);
            using IWorkbook workbook = FunctionalClasses.DataTableWriter.CreateWorkbook(format);
            FunctionalClasses.DataTableWriter.Write(workbook, table, options.WriteColumnNames);
            FunctionalClasses.DataTableWriter.Save(workbook, filePath);
        }

        /// <summary>
        /// Saves a single <see cref="DataTable"/> to a stream. <see cref="ExcelWriteOptions.Format"/> is required.
        /// The stream is not disposed.
        /// </summary>
        public static void Save(Stream stream, DataTable table, ExcelWriteOptions options)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(table);
            ArgumentNullException.ThrowIfNull(options);

            ExcelFormat format = ExcelFormatResolver.Resolve(options.Format, pathOrFileName: null);
            using IWorkbook workbook = FunctionalClasses.DataTableWriter.CreateWorkbook(format);
            FunctionalClasses.DataTableWriter.Write(workbook, table, options.WriteColumnNames);
            FunctionalClasses.DataTableWriter.Save(workbook, stream);
        }

        /// <summary>
        /// Saves a single <see cref="DataTable"/> to a stream with an explicit format.
        /// </summary>
        public static void Save(Stream stream, DataTable table, ExcelFormat format, bool writeColumnNames = true)
        {
            Save(stream, table, new ExcelWriteOptions
            {
                Format = format,
                WriteColumnNames = writeColumnNames,
            });
        }

        /// <summary>
        /// Saves every sheet from the book into an Excel file.
        /// </summary>
        public static void SaveBook(string filePath, XLBook book, params bool[] writeColumnNames)
        {
            SaveBook(filePath, book, new ExcelWriteOptions
            {
                WriteColumnNames = true,
                WriteColumnNamesPerSheet = writeColumnNames.Length > 0 ? writeColumnNames : null,
            });
        }

        /// <summary>
        /// Saves every sheet with the same header option for all sheets.
        /// </summary>
        public static void SaveBookParamsConnected(string filePath, XLBook book, bool writeColumnNames = true)
        {
            SaveBook(filePath, book, new ExcelWriteOptions { WriteColumnNames = writeColumnNames });
        }

        /// <summary>
        /// Saves every sheet from the book to a file using <see cref="ExcelWriteOptions"/>.
        /// </summary>
        public static void SaveBook(string filePath, XLBook book, ExcelWriteOptions options)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);
            ArgumentNullException.ThrowIfNull(book);
            ArgumentNullException.ThrowIfNull(options);

            ExcelFormat format = ExcelFormatResolver.Resolve(options.Format, filePath);
            using IWorkbook workbook = FunctionalClasses.DataTableWriter.CreateWorkbook(format);
            WriteBook(workbook, book, options);
            FunctionalClasses.DataTableWriter.Save(workbook, filePath);
        }

        /// <summary>
        /// Saves every sheet from the book to a stream. <see cref="ExcelWriteOptions.Format"/> is required.
        /// The stream is not disposed.
        /// </summary>
        public static void SaveBook(Stream stream, XLBook book, ExcelWriteOptions options)
        {
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(book);
            ArgumentNullException.ThrowIfNull(options);

            ExcelFormat format = ExcelFormatResolver.Resolve(options.Format, pathOrFileName: null);
            using IWorkbook workbook = FunctionalClasses.DataTableWriter.CreateWorkbook(format);
            WriteBook(workbook, book, options);
            FunctionalClasses.DataTableWriter.Save(workbook, stream);
        }

        /// <summary>
        /// Saves every sheet from the book to a stream with an explicit format.
        /// </summary>
        public static void SaveBook(Stream stream, XLBook book, ExcelFormat format, bool writeColumnNames = true)
        {
            SaveBook(stream, book, new ExcelWriteOptions
            {
                Format = format,
                WriteColumnNames = writeColumnNames,
            });
        }

        private static void WriteBook(IWorkbook workbook, XLBook book, ExcelWriteOptions options)
        {
            bool[] flags = new bool[book.Sheets.Count];
            for (int i = 0; i < book.Sheets.Count; i++)
            {
                flags[i] = options.WriteColumnNamesPerSheet is { Length: > 0 } perSheet && i < perSheet.Length
                    ? perSheet[i]
                    : options.WriteColumnNames;
            }

            FunctionalClasses.DataTableWriter.Write(workbook, book, flags);
        }
    }
}
