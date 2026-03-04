namespace OCRBulkAdd.Processing
{
    internal static class OcrTextNormalizer
    {
        /// <summary>
        /// make OCR output more readable:
        /// - unify line endings
        /// - remove empty lines
        /// - normalize different minus chars to 'real minus' ( - )
        /// </summary>
        public static string Normalize(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return string.Empty;

            s = s.Replace("\r\n", "\n").Replace('\r', '\n')
                 .Replace('−', '-')
                 .Replace('–', '-')
                 .Replace('—', '-')
                 .Replace('‐', '-');

            var lines = s.Split('\n')
                         .Select(l => l.Trim())
                         .Where(l => !string.IsNullOrWhiteSpace(l));

            return string.Join(Environment.NewLine, lines).Trim();
        }
    }
}