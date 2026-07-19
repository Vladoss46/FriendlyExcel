namespace FriendlyExcel
{
    /// <summary>
    /// Maps a POCO property to an Excel column name (e.g. Russian headers).
    /// When omitted, the property name is used.
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
    public sealed class ExcelColumnAttribute : Attribute
    {
        /// <summary>
        /// Excel column header name.
        /// </summary>
        public string Name { get; }

        /// <param name="name">Column header as it appears in Excel.</param>
        public ExcelColumnAttribute(string name)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(name);
            Name = name;
        }
    }
}
