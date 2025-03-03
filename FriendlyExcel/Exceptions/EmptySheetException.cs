using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FriendlyExcel.Exceptions
{
    internal class EmptySheetException : NullReferenceException
    {
        public EmptySheetException(string? message) : base(message)
        {
        }

        public EmptySheetException(string? message, Exception? innerException) : base(message, innerException)
        {
        }
    }
}
