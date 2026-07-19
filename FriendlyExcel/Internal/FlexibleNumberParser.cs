using System.Globalization;

namespace FriendlyExcel.Internal
{
    /// <summary>
    /// Parses numbers from Excel text cells where the decimal separator may be
    /// ',' (ru-RU / many EU locales) or '.' (invariant / en-US).
    /// </summary>
    /// <remarks>
    /// Russian Excel UI uses a comma as the decimal separator. Proper numeric cells
    /// are still stored as binary doubles (culture-agnostic). Text cells may contain
    /// "12,5" or "12.5". InvariantCulture must NOT be used blindly on "12,5" —
    /// there ',' is a thousands separator and the value becomes 125.
    /// </remarks>
    internal static class FlexibleNumberParser
    {
        private static readonly CultureInfo Russian = CultureInfo.GetCultureInfo("ru-RU");
        private static readonly CultureInfo German = CultureInfo.GetCultureInfo("de-DE");
        private static readonly CultureInfo UsEnglish = CultureInfo.GetCultureInfo("en-US");

        public static bool TryParseDouble(string? text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text))
                return false;

            string s = text.Trim();
            bool hasComma = s.Contains(',');
            bool hasDot = s.Contains('.');

            // "12,5" / "1 234,56" — decimal comma (Russian Excel text)
            if (hasComma && !hasDot)
            {
                return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, Russian, out value)
                    || double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, German, out value);
            }

            // "12.5" / "1,234.56" — decimal point
            if (hasDot && !hasComma)
            {
                return double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value)
                    || double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, UsEnglish, out value);
            }

            // Both separators: the last one is the decimal separator.
            // Normalize to invariant "1234.56" because ru-RU group separator is often
            // a narrow/nbsp space, not '.', so "1.234,56" would not parse with ru-RU alone.
            if (hasComma && hasDot)
            {
                int lastComma = s.LastIndexOf(',');
                int lastDot = s.LastIndexOf('.');
                string normalized = lastComma > lastDot
                    ? s.Replace(".", string.Empty).Replace(',', '.')   // 1.234,56 → 1234.56
                    : s.Replace(",", string.Empty);                    // 1,234.56 → 1234.56

                return double.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
            }

            // Integer-like text: "42"
            return double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || double.TryParse(s, NumberStyles.Float, Russian, out value);
        }

        public static bool TryParseInt32(string? text, out int value)
        {
            value = 0;
            if (!TryParseDouble(text, out double d))
                return false;

            if (d < int.MinValue || d > int.MaxValue || !double.IsInteger(d))
                return false;

            value = (int)d;
            return true;
        }
    }
}
