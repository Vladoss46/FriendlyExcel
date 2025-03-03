using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel.Extensions
{
    internal static class CellExtensions
    {
        public static string GetValueAsString(this ICell cell)
        {
            List<Func<string>> funcArray = [
                ()=> cell.StringCellValue,
                ()=> cell.NumericCellValue.ToString(),
                ()=> cell.BooleanCellValue.ToString(),
                ()=> cell.DateCellValue?.ConvertDateToString() ?? "",
                ];

            string value = RecursiveChecking(0, funcArray.ToArray());

            return value;
        }

        private static string RecursiveChecking(int index, params Func<string>[] func)
        {
            try
            {
                return func[index]();
            }
            catch (InvalidOperationException)
            {
                return RecursiveChecking(index + 1, func);
            }
        }
    }
}
