namespace FriendlyExcel
{
    /// <summary>
    /// Options for writing Excel files.
    /// </summary>
    public sealed class ExcelWriteOptions
    {
        /// <summary>
        /// When <c>true</c>, writes column names as the first row.
        /// Used for single-sheet saves and as the default for books.
        /// </summary>
        public bool WriteColumnNames { get; init; } = true;

        /// <summary>
        /// Per-sheet header flags for <see cref="Models.XLBook"/> writes.
        /// Missing entries default to <see cref="WriteColumnNames"/>.
        /// </summary>
        public bool[]? WriteColumnNamesPerSheet { get; init; }

        /// <summary>
        /// Output format. Required for <see cref="Stream"/> targets.
        /// When writing to a path, may be omitted and inferred from the extension.
        /// </summary>
        public ExcelFormat? Format { get; init; }
    }
}
