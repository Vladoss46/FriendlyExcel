// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using FriendlyExcel.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using System.Globalization;

namespace FriendlyExcel.FunctionalClasses
{
    internal static class DataTableWriter
    {
        private const int MaxSheetNameLength = 31;

        public static IWorkbook CreateWorkbook(string filePath)
        {
            return CreateWorkbook(Internal.ExcelFormatResolver.FromPathOrFileName(filePath));
        }

        public static IWorkbook CreateWorkbook(ExcelFormat format)
        {
            return format switch
            {
                ExcelFormat.Xls => new HSSFWorkbook(),
                ExcelFormat.Xlsx => new XSSFWorkbook(),
                _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported Excel format."),
            };
        }

        public static void WriteSheet(IWorkbook workbook, DataTable table, bool writeColumnNames = true)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            ArgumentNullException.ThrowIfNull(table);

            ISheet sheet = workbook.CreateSheet(ResolveSheetName(workbook, table.TableName));
            ICellStyle? dateTimeStyle = null;
            ICellStyle? dateOnlyStyle = null;
            ICellStyle? timeOnlyStyle = null;

            int rowIndex = 0;
            if (writeColumnNames && table.Columns.Count > 0)
            {
                IRow headerRow = sheet.CreateRow(rowIndex++);
                for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    headerRow.CreateCell(columnIndex).SetCellValue(table.Columns[columnIndex].ColumnName);
                }
            }

            foreach (DataRow dataRow in table.Rows)
            {
                IRow row = sheet.CreateRow(rowIndex++);
                for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    object value = dataRow[columnIndex];
                    if (value is DBNull || value is null)
                        continue;

                    ICell cell = row.CreateCell(columnIndex);
                    SetCellValue(
                        cell,
                        value,
                        table.Columns[columnIndex].DataType,
                        workbook,
                        ref dateTimeStyle,
                        ref dateOnlyStyle,
                        ref timeOnlyStyle);
                }
            }
        }

        public static void Write(IWorkbook workbook, DataTable table, bool writeColumnNames = true)
        {
            WriteSheet(workbook, table, writeColumnNames);
        }

        public static void Write(IWorkbook workbook, XLBook book, params bool[] writeColumnNames)
        {
            ArgumentNullException.ThrowIfNull(book);

            for (int i = 0; i < book.Sheets.Count; i++)
            {
                bool useColumnNames = i < writeColumnNames.Length ? writeColumnNames[i] : true;
                WriteSheet(workbook, book.Sheets[i], useColumnNames);
            }
        }

        public static void Save(IWorkbook workbook, string filePath)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            ArgumentException.ThrowIfNullOrWhiteSpace(filePath);

            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrEmpty(directory))
                Directory.CreateDirectory(directory);

            using FileStream fs = new(filePath, FileMode.Create, FileAccess.Write);
            Save(workbook, fs);
        }

        public static void Save(IWorkbook workbook, Stream stream)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            ArgumentNullException.ThrowIfNull(stream);

            if (!stream.CanWrite)
                throw new ArgumentException("Stream must be writable.", nameof(stream));

            workbook.Write(stream, leaveOpen: true);
            if (stream.CanSeek)
                stream.Flush();
        }

        private static void SetCellValue(
            ICell cell,
            object value,
            Type columnType,
            IWorkbook workbook,
            ref ICellStyle? dateTimeStyle,
            ref ICellStyle? dateOnlyStyle,
            ref ICellStyle? timeOnlyStyle)
        {
            Type type = Nullable.GetUnderlyingType(columnType) ?? columnType;

            if (type == typeof(string))
            {
                cell.SetCellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                return;
            }

            if (type == typeof(bool))
            {
                cell.SetCellValue(Convert.ToBoolean(value, CultureInfo.InvariantCulture));
                return;
            }

            if (type == typeof(int) || type == typeof(long) || type == typeof(short) || type == typeof(byte))
            {
                cell.SetCellValue(Convert.ToDouble(value, CultureInfo.InvariantCulture));
                return;
            }

            if (type == typeof(double) || type == typeof(float) || type == typeof(decimal))
            {
                cell.SetCellValue(Convert.ToDouble(value, CultureInfo.InvariantCulture));
                return;
            }

            if (type == typeof(DateTime))
            {
                dateTimeStyle ??= CreateDateStyle(workbook, "yyyy-MM-dd HH:mm:ss");
                cell.SetCellValue((DateTime)value);
                cell.CellStyle = dateTimeStyle;
                return;
            }

            if (type == typeof(DateOnly))
            {
                dateOnlyStyle ??= CreateDateStyle(workbook, "yyyy-MM-dd");
                cell.SetCellValue(((DateOnly)value).ToDateTime(TimeOnly.MinValue));
                cell.CellStyle = dateOnlyStyle;
                return;
            }

            if (type == typeof(TimeOnly))
            {
                timeOnlyStyle ??= CreateDateStyle(workbook, "HH:mm:ss");
                cell.SetCellValue(DateTime.FromOADate(((TimeOnly)value).ToTimeSpan().TotalDays));
                cell.CellStyle = timeOnlyStyle;
                return;
            }

            cell.SetCellValue(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
        }

        private static ICellStyle CreateDateStyle(IWorkbook workbook, string format)
        {
            ICellStyle style = workbook.CreateCellStyle();
            style.DataFormat = workbook.CreateDataFormat().GetFormat(format);
            return style;
        }

        private static string ResolveSheetName(IWorkbook workbook, string? tableName)
        {
            string baseName = string.IsNullOrWhiteSpace(tableName)
                ? $"Sheet{workbook.NumberOfSheets + 1}"
                : SanitizeSheetName(tableName);

            if (workbook.GetSheet(baseName) is null)
                return baseName;

            int suffix = 2;
            while (true)
            {
                string candidate = TruncateSheetName($"{baseName}_{suffix}");
                if (workbook.GetSheet(candidate) is null)
                    return candidate;
                suffix++;
            }
        }

        private static string SanitizeSheetName(string name)
        {
            char[] invalid = ['\\', '/', '?', '*', '[', ']', ':'];
            string sanitized = name;
            foreach (char c in invalid)
                sanitized = sanitized.Replace(c, '_');

            sanitized = sanitized.Trim();
            if (string.IsNullOrEmpty(sanitized))
                sanitized = "Sheet1";

            return TruncateSheetName(sanitized);
        }

        private static string TruncateSheetName(string name)
        {
            return name.Length <= MaxSheetNameLength
                ? name
                : name[..MaxSheetNameLength];
        }

    }
}
