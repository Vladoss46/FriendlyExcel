// This project is licensed under the MIT License.

using System.Globalization;

namespace FriendlyExcel.Extensions
{
    internal static class CommonExtensions
    {
        public static string ConvertDateToString(this DateTime date)
        {
            return date.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }
    }
}
