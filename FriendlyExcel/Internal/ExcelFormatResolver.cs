namespace FriendlyExcel.Internal
{
    internal static class ExcelFormatResolver
    {
        public static ExcelFormat FromPathOrFileName(string pathOrFileName)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(pathOrFileName);

            if (pathOrFileName.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                return ExcelFormat.Xlsx;

            if (pathOrFileName.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                return ExcelFormat.Xls;

            throw new ArgumentException(
                $"Cannot infer Excel format from '{pathOrFileName}'. Use .xls / .xlsx or specify ExcelFormat explicitly.",
                nameof(pathOrFileName));
        }

        public static ExcelFormat Resolve(ExcelFormat? format, string? pathOrFileName)
        {
            if (format.HasValue)
                return format.Value;

            if (!string.IsNullOrWhiteSpace(pathOrFileName))
                return FromPathOrFileName(pathOrFileName);

            throw new ArgumentException(
                "Excel format must be specified when the file name/extension is unknown. Set Format or FileName.");
        }
    }
}
