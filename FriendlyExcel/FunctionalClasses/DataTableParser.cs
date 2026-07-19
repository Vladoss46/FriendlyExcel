// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using FriendlyExcel.Exceptions;
using FriendlyExcel.Extensions;
using FriendlyExcel.Internal;
using NPOI.SS.UserModel;
using System.Data;
using System.Globalization;

namespace FriendlyExcel.FunctionalClasses
{
    internal static class DataTableParser
    {
        public static DataTable Parse(ISheet sheet, bool useFirstRowAsColumnNames = true)
        {
            string[] columnNames = useFirstRowAsColumnNames ? GetColumnNames(sheet) : GetBaseColumnNames(sheet);
            int startDataRowIndex = useFirstRowAsColumnNames ? 1 : 0;
            Type[] columnTypes = GetColumnTypes(sheet, startDataRowIndex);
            DataTable table = CreateDataTable(columnNames, columnTypes);
            table.TableName = sheet.SheetName;
            return FillTable(table, sheet, columnTypes, startDataRowIndex);
        }

        public static DataTable Parse(ISheet sheet, Type[] columnTypes, bool useFirstRowAsColumnNames = true)
        {
            if (columnTypes.Length == 0)
                throw new ArgumentException("columnTypes must contain at least one type.", nameof(columnTypes));

            string[] columnNames = useFirstRowAsColumnNames ? GetColumnNames(sheet) : GetBaseColumnNames(sheet);
            int startDataRowIndex = useFirstRowAsColumnNames ? 1 : 0;
            DataTable table = CreateDataTable(columnNames, columnTypes);
            table.TableName = sheet.SheetName;
            return FillTable(table, sheet, columnTypes, startDataRowIndex);
        }

        internal static string[] GetColumnNames(ISheet sheet)
        {
            IRow firstRow = sheet.GetRow(sheet.FirstRowNum) ?? throw new EmptySheetException("Not found any data at the sheet");
            List<string> columnNames = [];
            foreach (ICell cell in firstRow.Cells)
            {
                string columnName = cell.GetValueAsString();
                if (string.IsNullOrWhiteSpace(columnName))
                    throw new ArgumentException($"ColumnName at {cell.Address.FormatAsString()} position is empty");
                columnNames.Add(columnName);
            }
            return [.. columnNames];
        }

        internal static string[] GetBaseColumnNames(ISheet sheet)
        {
            IRow firstRow = sheet.GetRow(sheet.FirstRowNum) ?? throw new EmptySheetException("Not found any data at the sheet");
            List<string> columnNames = [];
            for (int index = 0; index < firstRow.Cells.Count; index++)
                columnNames.Add($"Column{index + 1}");
            return [.. columnNames];
        }

        internal static Type[] GetColumnTypes(ISheet sheet, int firstRowIndex)
        {
            int index = sheet.FirstRowNum + firstRowIndex;
            IRow? currentRow = sheet.GetRow(index++);
            if (currentRow is null)
                throw new EmptySheetException("Not found any data at the sheet");

            Type[] columnTypes = [];
            do
            {
                List<Type> currentColumnTypes = [];
                foreach (ICell cell in currentRow.Cells)
                    currentColumnTypes.Add(ResolveCellType(cell));

                columnTypes = CompareAndSwitch([.. currentColumnTypes], columnTypes);
                currentRow = sheet.GetRow(index++);
            }
            while (currentRow is not null);

            return columnTypes;
        }

        internal static Type GetTypeOfCell(CellType cellType)
        {
            return cellType switch
            {
                CellType.Boolean => typeof(bool),
                CellType.String => typeof(string),
                CellType.Numeric => typeof(double),
                CellType.Formula => typeof(string),
                _ => typeof(string),
            };
        }

        internal static Type CheckDoubleType(double value)
        {
            if (!double.IsNaN(value)
                && !double.IsInfinity(value)
                && value >= int.MinValue
                && value <= int.MaxValue
                && double.IsInteger(value))
            {
                return typeof(int);
            }

            return typeof(double);
        }

        internal static Type CheckDateTimeType(ICell cell)
        {
            if (!DateUtil.IsCellDateFormatted(cell))
                return typeof(double);

            double numericCellValue = cell.NumericCellValue;

            if (numericCellValue < 1.0d)
                return typeof(TimeOnly);

            if (double.IsInteger(numericCellValue))
                return typeof(DateOnly);

            return typeof(DateTime);
        }

        internal static Type[] CompareAndSwitch(Type[] fromArray, Type[] toArray)
        {
            if (toArray.Length == 0)
                return fromArray;

            Type[] result = [.. toArray];
            for (int index = 0; index < fromArray.Length && index < result.Length; index++)
            {
                Type current = fromArray[index];
                Type accumulated = result[index];

                if (current == typeof(string) || accumulated == typeof(string))
                {
                    result[index] = typeof(string);
                    continue;
                }

                if ((current == typeof(double) && accumulated == typeof(int))
                    || (current == typeof(int) && accumulated == typeof(double)))
                {
                    result[index] = typeof(double);
                }
            }

            return result;
        }

        internal static DataTable CreateDataTable(string[] columnNames, Type[] columnTypes)
        {
            if (columnNames.Length != columnTypes.Length)
                throw new ArgumentException("Column names count doesn't match column types count.");

            DataTable table = new();
            for (int i = 0; i < columnNames.Length; i++)
                table.Columns.Add(columnNames[i], columnTypes[i]);
            return table;
        }

        internal static DataTable FillTable(DataTable table, ISheet sheet, Type[] columnTypes, int startRowIndex)
        {
            for (int i = sheet.FirstRowNum + startRowIndex; i <= sheet.LastRowNum; i++)
            {
                IRow? row = sheet.GetRow(i);
                if (row is null)
                    continue;

                DataRow tableRow = table.NewRow();
                ParseRow(row, tableRow, columnTypes);
                table.Rows.Add(tableRow);
            }

            table.AcceptChanges();
            return table;
        }

        internal static DataRow ParseRow(IRow row, DataRow resultRow, Type[] columnTypes)
        {
            for (int i = 0; i < row.Cells.Count && i < columnTypes.Length; i++)
            {
                ICell cell = row.Cells[i];
                resultRow[i] = GetCellValue(cell, columnTypes[i]);
            }
            return resultRow;
        }

        internal static object GetCellValue(ICell cell, Type type)
        {
            CellType effectiveType = GetEffectiveCellType(cell);
            if (effectiveType == CellType.Blank)
                return DBNull.Value;

            return type.Name switch
            {
                nameof(String) => cell.GetValueAsString(),
                nameof(Double) => ReadAsDouble(cell, effectiveType),
                nameof(Int32) => ReadAsInt32(cell, effectiveType),
                nameof(Boolean) => cell.BooleanCellValue,
                nameof(DateTime) => cell.DateCellValue!,
                nameof(TimeOnly) => cell.TimeOnlyCellValue!,
                nameof(DateOnly) => cell.DateOnlyCellValue!,
                _ => throw new NotSupportedException($"Unknown type of a column - {type.Name}"),
            };
        }

        private static Type ResolveCellType(ICell cell)
        {
            CellType effectiveType = GetEffectiveCellType(cell);

            // Textual numbers from Russian Excel often look like "12,5"
            if (effectiveType == CellType.String)
            {
                string text = cell.StringCellValue;
                if (FlexibleNumberParser.TryParseDouble(text, out double parsed))
                    return CheckDoubleType(parsed);
                return typeof(string);
            }

            Type type = GetTypeOfCell(effectiveType);
            if (type != typeof(double))
                return type;

            if (DateUtil.IsCellDateFormatted(cell))
                return CheckDateTimeType(cell);

            return CheckDoubleType(cell.NumericCellValue);
        }

        private static double ReadAsDouble(ICell cell, CellType effectiveType)
        {
            if (effectiveType is CellType.Numeric or CellType.Formula)
                return cell.NumericCellValue;

            if (effectiveType == CellType.String
                && FlexibleNumberParser.TryParseDouble(cell.StringCellValue, out double parsed))
            {
                return parsed;
            }

            throw new InvalidCastException(
                $"Cannot read cell {cell.Address} as Double (type {effectiveType}, text '{cell.GetValueAsString()}').");
        }

        private static int ReadAsInt32(ICell cell, CellType effectiveType)
        {
            if (effectiveType is CellType.Numeric or CellType.Formula)
                return Convert.ToInt32(cell.NumericCellValue, CultureInfo.InvariantCulture);

            if (effectiveType == CellType.String
                && FlexibleNumberParser.TryParseInt32(cell.StringCellValue, out int parsed))
            {
                return parsed;
            }

            throw new InvalidCastException(
                $"Cannot read cell {cell.Address} as Int32 (type {effectiveType}, text '{cell.GetValueAsString()}').");
        }

        private static CellType GetEffectiveCellType(ICell cell)
        {
            return cell.CellType == CellType.Formula
                ? cell.CachedFormulaResultType
                : cell.CellType;
        }
    }
}
