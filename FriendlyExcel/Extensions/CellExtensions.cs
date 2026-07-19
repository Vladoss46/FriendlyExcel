// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using NPOI.SS.UserModel;
using System.Globalization;

namespace FriendlyExcel.Extensions
{
    internal static class CellExtensions
    {
        private static readonly DataFormatter InvariantFormatter = new(CultureInfo.InvariantCulture);

        /// <summary>
        /// Returns a culture-stable string for any cell.
        /// Text (including Cyrillic) is preserved; numbers/dates use invariant formatting
        /// so results do not depend on the OS locale (e.g. ru-RU decimal comma).
        /// </summary>
        public static string GetValueAsString(this ICell cell)
        {
            if (cell is null || cell.CellType == CellType.Blank)
                return string.Empty;

            CellType type = cell.CellType == CellType.Formula
                ? cell.CachedFormulaResultType
                : cell.CellType;

            return type switch
            {
                CellType.String => cell.StringCellValue ?? string.Empty,
                CellType.Boolean => cell.BooleanCellValue ? "TRUE" : "FALSE",
                CellType.Numeric when DateUtil.IsCellDateFormatted(cell)
                    => FormatDateCell(cell),
                CellType.Numeric
                    => cell.NumericCellValue.ToString(CultureInfo.InvariantCulture),
                CellType.Error
                    => FormulaError.ForInt(cell.ErrorCellValue).String,
                _ => InvariantFormatter.FormatCellValue(cell) ?? string.Empty,
            };
        }

        private static string FormatDateCell(ICell cell)
        {
            DateTime? date = cell.DateCellValue;
            return date?.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) ?? string.Empty;
        }
    }
}
