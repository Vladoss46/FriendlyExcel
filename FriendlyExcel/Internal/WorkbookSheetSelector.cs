using NPOI.SS.UserModel;

namespace FriendlyExcel.Internal
{
    internal static class WorkbookSheetSelector
    {
        public static ISheet GetSheet(IWorkbook workbook, ExcelReadOptions options)
        {
            ArgumentNullException.ThrowIfNull(workbook);
            ArgumentNullException.ThrowIfNull(options);

            if (!string.IsNullOrWhiteSpace(options.SheetName))
                return GetSheetByName(workbook, options.SheetName);

            if (options.SheetIndex < 0 || options.SheetIndex >= workbook.NumberOfSheets)
            {
                throw new ArgumentOutOfRangeException(
                    nameof(options),
                    options.SheetIndex,
                    $"Sheet index {options.SheetIndex} is out of range (sheet count: {workbook.NumberOfSheets}).");
            }

            return workbook.GetSheetAt(options.SheetIndex);
        }

        public static ISheet GetSheetByName(IWorkbook workbook, string sheetName)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(sheetName);

            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                if (string.Equals(workbook.GetSheetName(i), sheetName, StringComparison.OrdinalIgnoreCase))
                    return workbook.GetSheetAt(i);
            }

            throw new ArgumentException($"Sheet '{sheetName}' was not found.", nameof(sheetName));
        }
    }
}
