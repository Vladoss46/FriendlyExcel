// This project is licensed under the MIT License.
// This file contains code from NPOI, which is licensed under the Apache License 2.0.

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace FriendlyExcel.FunctionalClasses
{
    internal static class DataTableReader
    {
        public static IWorkbook Read(string filePath)
        {
            ExcelFormat format = Internal.ExcelFormatResolver.FromPathOrFileName(filePath);
            using FileStream fs = new(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            return Read(fs, format);
        }

        public static IWorkbook Read(Stream stream, ExcelFormat format)
        {
            ArgumentNullException.ThrowIfNull(stream);

            if (!stream.CanRead)
                throw new ArgumentException("Stream must be readable.", nameof(stream));

            return format switch
            {
                ExcelFormat.Xls => new HSSFWorkbook(stream),
                ExcelFormat.Xlsx => new XSSFWorkbook(stream),
                _ => throw new ArgumentOutOfRangeException(nameof(format), format, "Unsupported Excel format."),
            };
        }
    }
}
