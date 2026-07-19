namespace FriendlyExcel
{
    /// <summary>
    /// Options for reading a single worksheet.
    /// </summary>
    public sealed class ExcelReadOptions
    {
        /// <summary>
        /// Zero-based sheet index. Used when <see cref="SheetName"/> is not set.
        /// </summary>
        public int SheetIndex { get; init; }

        /// <summary>
        /// Sheet name (case-insensitive). When set, takes priority over <see cref="SheetIndex"/>.
        /// </summary>
        public string? SheetName { get; init; }

        /// <summary>
        /// When <c>true</c>, the first row becomes column names.
        /// </summary>
        public bool UseFirstRowAsColumnNames { get; init; } = true;

        /// <summary>
        /// Optional explicit column types. When null, types are inferred.
        /// </summary>
        public Type[]? ColumnTypes { get; init; }

        /// <summary>
        /// Required when reading from a <see cref="Stream"/> without a file name.
        /// When reading from a path, may be omitted and inferred from the extension.
        /// </summary>
        public ExcelFormat? Format { get; init; }

        /// <summary>
        /// Optional file name (e.g. upload file name) used only to infer <see cref="Format"/>.
        /// </summary>
        public string? FileName { get; init; }
    }

    /// <summary>
    /// Options for reading an entire workbook.
    /// </summary>
    public sealed class ExcelReadBookOptions
    {
        /// <summary>
        /// Header flag applied to every sheet when <see cref="UseFirstRowAsColumnNamesPerSheet"/> is null.
        /// </summary>
        public bool UseFirstRowAsColumnNames { get; init; } = true;

        /// <summary>
        /// Per-sheet header flags. Missing entries default to <see cref="UseFirstRowAsColumnNames"/>.
        /// </summary>
        public bool[]? UseFirstRowAsColumnNamesPerSheet { get; init; }

        /// <summary>
        /// Required when reading from a <see cref="Stream"/> without a file name.
        /// </summary>
        public ExcelFormat? Format { get; init; }

        /// <summary>
        /// Optional file name used only to infer <see cref="Format"/>.
        /// </summary>
        public string? FileName { get; init; }
    }
}
