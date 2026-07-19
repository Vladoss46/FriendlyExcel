using System.Collections.Concurrent;
using System.Data;
using System.Globalization;
using System.Reflection;

namespace FriendlyExcel.Internal
{
    internal static class PocoMapper
    {
        private static readonly ConcurrentDictionary<Type, PropertyMap[]> Maps = new();

        public static List<T> FromTable<T>(DataTable table) where T : new()
        {
            ArgumentNullException.ThrowIfNull(table);

            PropertyMap[] maps = GetMaps(typeof(T));
            Dictionary<string, string> resolvedColumns = ResolveColumns(table, maps, typeof(T));

            List<T> result = new(table.Rows.Count);
            for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
            {
                DataRow row = table.Rows[rowIndex];
                T item = new();
                foreach (PropertyMap map in maps)
                {
                    string actualColumn = resolvedColumns[map.ColumnName];
                    object raw = row[actualColumn];
                    try
                    {
                        object? converted = ConvertValue(raw, map.Property.PropertyType);
                        map.Property.SetValue(item, converted);
                    }
                    catch (Exception ex) when (ex is not InvalidOperationException)
                    {
                        throw new InvalidOperationException(
                            $"Cannot map column '{map.ColumnName}' to {typeof(T).Name}.{map.Property.Name} at data row {rowIndex}.",
                            ex);
                    }
                }

                result.Add(item);
            }

            return result;
        }

        public static DataTable ToTable<T>(IEnumerable<T> rows, string? tableName = null)
        {
            ArgumentNullException.ThrowIfNull(rows);

            PropertyMap[] maps = GetMaps(typeof(T));
            if (maps.Length == 0)
            {
                throw new InvalidOperationException(
                    $"Type {typeof(T).Name} has no public get/set instance properties to map.");
            }

            DataTable table = new(string.IsNullOrWhiteSpace(tableName) ? typeof(T).Name : tableName);
            foreach (PropertyMap map in maps)
            {
                Type columnType = Nullable.GetUnderlyingType(map.Property.PropertyType) ?? map.Property.PropertyType;
                table.Columns.Add(map.ColumnName, columnType);
            }

            foreach (T? item in rows)
            {
                if (item is null)
                    continue;

                DataRow row = table.NewRow();
                for (int i = 0; i < maps.Length; i++)
                {
                    object? value = maps[i].Property.GetValue(item);
                    row[i] = value ?? DBNull.Value;
                }

                table.Rows.Add(row);
            }

            table.AcceptChanges();
            return table;
        }

        private static PropertyMap[] GetMaps(Type type)
        {
            return Maps.GetOrAdd(type, static t =>
            {
                List<PropertyMap> maps = [];
                foreach (PropertyInfo property in t.GetProperties(BindingFlags.Instance | BindingFlags.Public))
                {
                    if (!property.CanRead || !property.CanWrite || property.GetIndexParameters().Length > 0)
                        continue;

                    ExcelColumnAttribute? attr = property.GetCustomAttribute<ExcelColumnAttribute>();
                    string columnName = attr?.Name ?? property.Name;
                    maps.Add(new PropertyMap(property, columnName));
                }

                return [.. maps];
            });
        }

        /// <summary>
        /// Maps requested column names (case-insensitive) to actual DataTable column names.
        /// </summary>
        private static Dictionary<string, string> ResolveColumns(DataTable table, PropertyMap[] maps, Type type)
        {
            Dictionary<string, string> byIgnoreCase = new(StringComparer.OrdinalIgnoreCase);
            foreach (DataColumn column in table.Columns)
                byIgnoreCase[column.ColumnName] = column.ColumnName;

            Dictionary<string, string> resolved = new(StringComparer.OrdinalIgnoreCase);
            foreach (PropertyMap map in maps)
            {
                if (!byIgnoreCase.TryGetValue(map.ColumnName, out string? actual))
                {
                    throw new InvalidOperationException(
                        $"Column '{map.ColumnName}' required by {type.Name}.{map.Property.Name} was not found in the sheet.");
                }

                resolved[map.ColumnName] = actual;
            }

            return resolved;
        }

        private static object? ConvertValue(object raw, Type targetType)
        {
            if (raw is DBNull || raw is null)
            {
                Type? underlying = Nullable.GetUnderlyingType(targetType);
                if (underlying is not null || !targetType.IsValueType)
                    return null;

                return Activator.CreateInstance(targetType);
            }

            Type nonNullType = Nullable.GetUnderlyingType(targetType) ?? targetType;
            if (nonNullType.IsInstanceOfType(raw))
                return raw;

            if (nonNullType == typeof(string))
                return Convert.ToString(raw, CultureInfo.InvariantCulture);

            if (raw is string text)
            {
                if (nonNullType == typeof(bool))
                {
                    if (bool.TryParse(text, out bool b))
                        return b;
                    if (string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(text, "FALSE", StringComparison.OrdinalIgnoreCase))
                    {
                        return string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase);
                    }
                }

                if (IsNumeric(nonNullType) && FlexibleNumberParser.TryParseDouble(text, out double parsed))
                    return ConvertFromDouble(parsed, nonNullType);

                if (nonNullType == typeof(DateTime) && DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.RoundtripKind, out DateTime dt))
                    return dt;

                if (nonNullType == typeof(DateOnly) && DateOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly d))
                    return d;

                if (nonNullType == typeof(TimeOnly) && TimeOnly.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out TimeOnly t))
                    return t;
            }

            if (raw is IConvertible && IsNumeric(nonNullType))
            {
                double d = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return ConvertFromDouble(d, nonNullType);
            }

            if (nonNullType == typeof(DateOnly) && raw is DateTime dateTimeForDate)
                return DateOnly.FromDateTime(dateTimeForDate);

            if (nonNullType == typeof(TimeOnly) && raw is DateTime dateTimeForTime)
                return TimeOnly.FromDateTime(dateTimeForTime);

            if (nonNullType == typeof(DateTime) && raw is DateOnly dateOnly)
                return dateOnly.ToDateTime(TimeOnly.MinValue);

            if (nonNullType == typeof(bool))
                return Convert.ToBoolean(raw, CultureInfo.InvariantCulture);

            return Convert.ChangeType(raw, nonNullType, CultureInfo.InvariantCulture);
        }

        private static bool IsNumeric(Type type)
        {
            return type == typeof(int)
                || type == typeof(long)
                || type == typeof(short)
                || type == typeof(byte)
                || type == typeof(double)
                || type == typeof(float)
                || type == typeof(decimal);
        }

        private static object ConvertFromDouble(double value, Type targetType)
        {
            if (targetType == typeof(double))
                return value;
            if (targetType == typeof(float))
                return (float)value;
            if (targetType == typeof(decimal))
                return (decimal)value;
            if (targetType == typeof(int))
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            if (targetType == typeof(long))
                return Convert.ToInt64(value, CultureInfo.InvariantCulture);
            if (targetType == typeof(short))
                return Convert.ToInt16(value, CultureInfo.InvariantCulture);
            if (targetType == typeof(byte))
                return Convert.ToByte(value, CultureInfo.InvariantCulture);

            return Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }

        private sealed record PropertyMap(PropertyInfo Property, string ColumnName);
    }
}
