// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Metrics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel.FunctionalClasses
{
    internal class DataTableParser
    {
        public static DataTable Parse(ISheet sheet, bool useFirstRowAsColumnNames = true)
        {
            string[] columnNames = useFirstRowAsColumnNames ? GetColumnNames(sheet) : GetBaseColumnNames(sheet);
            int startDataRowIndex = useFirstRowAsColumnNames ? 1 : 0;
            Type[] columnTypes = GetColumnTypes(sheet, startDataRowIndex);
            DataTable table = CreateDataTable(columnNames, columnTypes);
            table = FillTable(table, sheet, columnTypes, startDataRowIndex);
            return table;
        }
        private static string[] GetColumnNames(ISheet sheet)
        {
            IRow firstRow = sheet.GetRow(0) ?? throw new NullReferenceException("First row is empty");
            List<string> columnNames = [];
            foreach (var cell in firstRow.Cells)
            {
                string columnName = cell.StringCellValue;
                if (string.IsNullOrWhiteSpace(columnName))
                    throw new NullReferenceException($"ColumnName at {cell.Address.FormatAsString()} position is empty");//TODO: проверь, как FormatAsString работает
                columnNames.Add(columnName);
            }
            return [.. columnNames];
        }
        private static string[] GetBaseColumnNames(ISheet sheet)
        {
            IRow firstRow = sheet.GetRow(0) ?? throw new NullReferenceException("First row is empty");
            List<string> columnNames = [];
            foreach (var column in (sheet.GetRow(0).Cells.Select((_, index) => new { index })))
            {
                columnNames.Add($"Column{column.index + 1}");
            }
            return [.. columnNames];
        }
        private static Type[] GetColumnTypes(ISheet sheet, int firstRowIndex)
        {
            int index = firstRowIndex;
            IRow currentRow = sheet.GetRow(index++);
            Type[] columnTypes = [];
            do
            {
                List<Type> currentColumnTypes = [];

                foreach (var cell in currentRow.Cells)
                {
                    Type type = GetTypeOfCell(cell.CellType);
                    if (type == typeof(double))
                    {
                        type = CheckDoubleType(cell.NumericCellValue.ToString());
                    }
                    currentColumnTypes.Add(type);
                }
                columnTypes = CompareAndSwitch(currentColumnTypes.ToArray(), columnTypes.ToArray());


                currentRow = sheet.GetRow(index++);
            }
            while (currentRow is not null);
            return columnTypes;

            Type GetTypeOfCell(CellType cellType)
            {
                switch (cellType)
                {
                    default: return typeof(string);
                    case CellType.Boolean: return typeof(bool);
                    case CellType.String: return typeof(string);
                    case CellType.Numeric: return typeof(double);
                }
            }
            Type CheckDoubleType(string stringValue)
            {
                if (int.TryParse(stringValue, out _))
                {
                    return typeof(int);
                }
                return typeof(double);
            }
            Type[] CompareAndSwitch(Type[] fromArray, Type[] toArray)
            {
                //i try to find 
                if (toArray.Length == 0) return fromArray;

                Type[] toArray_Final = toArray;
                foreach (var item in fromArray.Select((value, index) => new { value, index }))
                {
                    if (item.value == typeof(string) &&
                        toArray_Final[item.index] != typeof(string))
                    {
                        toArray_Final[item.index] = item.value;
                    }
                }
                return toArray_Final;
            }
        }
        private static DataTable CreateDataTable(string[] columnNames, Type[] columnTypes)
        {
            DataTable table = new();
            if (columnNames.Length != columnTypes.Length)
            {
                throw new ArgumentException("Column names count doesn't compare to column types count. Send this error to author and show the case");
            }
            for (int i = 0; i < columnNames.Length; i++)
            {
                table.Columns.Add(columnNames[i], columnTypes[i]);
            }
            return table;
        }
        private static DataTable FillTable(DataTable table, ISheet sheet, Type[] columnTypes, int startRowIndex)//TODO: Я остановился перед FillTable методом в написании тестов
        {
            for (int i = startRowIndex; i < sheet.Count(); i++)
            {
                DataRow tableRow = table.NewRow();
                tableRow = ParseRow(sheet.GetRow(i), tableRow, columnTypes);
                tableRow.AcceptChanges();
            }
            return table;

            static DataRow ParseRow(IRow row, DataRow resultRow, Type[] columnTypes)
            {
                for (int i = 0; i < row.Cells.Count; i++)
                {
                    ICell cell = row.Cells[i];
                    resultRow[i] = GetCellValue(cell, columnTypes[i]);
                }
                return resultRow; //TODO:  Ну и это проверь, как работает, что-ли

                static object GetCellValue(ICell cell, Type type)
                {
                    return type.Name switch
                    {
                        "String" => cell.StringCellValue,
                        "Double" => cell.NumericCellValue,
                        "Int32" => Int32.Parse(cell.NumericCellValue.ToString()),
                        "Boolean" => cell.BooleanCellValue,
                        _ => throw new Exception($"Unknown type of a column - {type.Name}"),
                    };
                }
            }
        }
    }
}
