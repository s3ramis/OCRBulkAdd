using System.Globalization;
using System.Text.RegularExpressions;

namespace OCRBulkAdd.Processing
{
    internal static class NumberExtractor
    {
        // regex string for number-like tokens
        private static readonly Regex NumberTokenRegex = new Regex(
            @"(?:(?:[+\-]|−|–|—)[ \t\u00A0\u202F]*)?(?:(?:\d{1,3}(?:[., \t\u00A0\u202F]\d{3})+|\d+)(?:[.,]\d+)?|[.,]\d+)",
            RegexOptions.Compiled);

        // most numbers will be a currency
        private const int MaxDecimalDigits = 2;

        public static (decimal Sum, int Count) SumFromText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return (0m, 0);

            decimal sum = 0m;
            int count = 0;

            foreach (Match m in NumberTokenRegex.Matches(text))
            {
                if (TryParseByLastSeparatorRule(m.Value, out var value))
                {
                    sum += value;
                    count++;
                }
            }

            return (sum, count);
        }

        public static string ExtractNormalizedNumbersText(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return string.Empty;

            var numbers = new List<string>();

            foreach (Match m in NumberTokenRegex.Matches(text))
            {
                if (TryParseByLastSeparatorRule(m.Value, out var value))
                {
                    numbers.Add(value.ToString("0.00", CultureInfo.InvariantCulture));
                }
            }

            return string.Join(Environment.NewLine, numbers);
        }

        /// <summary>
        /// OCR often confuses . and , (thousands vs decimal seperator)
        /// we interpret the last . or , as the decimal seperator if <= 2 digits follow
        /// remove all other separators to get value that can easily be parsed into a decimal
        /// </summary>
        public static bool TryParseByLastSeparatorRule(string raw, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            // normalize minusses into the real deal
            string s = raw.Trim()
                .Replace('−', '-')
                .Replace('–', '-')
                .Replace('—', '-')
                .Replace('‐', '-');

            // retard acouting negative (12345) really means -12345
            bool negParen = s.StartsWith("(") && s.EndsWith(")");
            if (negParen)
                s = s.Substring(1, s.Length - 2);

            // remove whitespaces
            s = s.Replace(" ", "")
                 .Replace("\t", "")
                 .Replace("\u00A0", "")
                 .Replace("\u202F", "");

            // retarded account part 2: evaluate if number is negative or positive
            bool negative = negParen;
            if (s.StartsWith("+"))
                s = s.Substring(1);
            else if (s.StartsWith("-"))
            {
                negative = true;
                s = s.Substring(1);
            }

            // find last seperator
            int lastDot = s.LastIndexOf('.');
            int lastComma = s.LastIndexOf(',');
            int lastSepIndex = Math.Max(lastDot, lastComma);

            int digitsAfterSep = 0;
            if (lastSepIndex >= 0)
                digitsAfterSep = s.Length - lastSepIndex - 1;

            // keep only digits
            string digitsOnly = new string(s.Where(char.IsDigit).ToArray());
            if (digitsOnly.Length == 0)
                return false;

            string normalized;

            // only keep last seperator if it looks to be a money number
            if (lastSepIndex >= 0 && digitsAfterSep > 0 && digitsAfterSep <= MaxDecimalDigits)
            {
                // add leading zero for decimal-only numbers
                if (digitsOnly.Length <= digitsAfterSep)
                    digitsOnly = digitsOnly.PadLeft(digitsAfterSep + 1, '0');

                normalized = digitsOnly.Insert(digitsOnly.Length - digitsAfterSep, ".");
            }
            else
            {
                normalized = digitsOnly;
            }

            if (!decimal.TryParse(normalized,
                    NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture,
                    out var parsed))
                return false;

            value = negative ? -parsed : parsed;
            return true;
        }
    }
}