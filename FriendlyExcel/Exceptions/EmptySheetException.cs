namespace FriendlyExcel.Exceptions
{
    /// <summary>
    /// Thrown when a worksheet contains no usable rows.
    /// </summary>
    public class EmptySheetException : Exception
    {
        public EmptySheetException(string? message) : base(message)
        {
        }

        public EmptySheetException(string? message, Exception? innerException) : base(message, innerException)
        {
        }
    }
}
