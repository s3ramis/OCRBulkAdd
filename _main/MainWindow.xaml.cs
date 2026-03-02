using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Input;
using TesseractOCR;
using TesseractOCR.Enums;

using OcrLanguage = TesseractOCR.Enums.Language;

namespace OCRBulkAdd
{
    public partial class MainWindow : Window
    {
        private byte[] _currentImageBytes = Array.Empty<byte>();

        private readonly SemaphoreSlim _ocrLock = new SemaphoreSlim(1, 1);
        private const int MaxDecimalDigits = 2; 
        private static readonly Regex NumberTokenRegex = new Regex(@"(?:(?:[+\-]|−|–|—)[ \t\u00A0\u202F]*)?(?:(?:\d{1,3}(?:[., \t\u00A0\u202F]\d{3})+|\d+)(?:[.,]\d+)?|[.,]\d+)", RegexOptions.Compiled);

        private static readonly HashSet<string> ImageExtensions =
            new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            { ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".tif", ".tiff", ".webp" };

        private static readonly string TessdataPath = ResolveTessdataPath();

        public MainWindow()
        {
            InitializeComponent();
            UpdateHintVisibility();
        }

        private static string ResolveTessdataPath()
        {
            string baseDir = AppContext.BaseDirectory;

            string candidate = Path.Combine(baseDir, "tessdata");
            if (Directory.Exists(candidate)) return candidate;

            candidate = Path.Combine(Directory.GetCurrentDirectory(), "tessdata");
            if (Directory.Exists(candidate)) return candidate;

            candidate = Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "tessdata"));
            if (Directory.Exists(candidate)) return candidate;

            return Path.Combine(baseDir, "tessdata");
        }

        private async void Window_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key != Key.V) return;
            if ((Keyboard.Modifiers & ModifierKeys.Control) != ModifierKeys.Control) return;

            if (TryGetClipboardImageBytes(out var bytes) || TryGetClipboardImageFileBytes(out bytes))
            {
                SetImageBytes(bytes);
                e.Handled = true;
                await RunOcrAsync();
            }
        }

        private static bool TryGetClipboardImageBytes(out byte[] bytes)
        {
            bytes = Array.Empty<byte>();

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (Clipboard.ContainsImage())
                    {
                        var bmp = Clipboard.GetImage();
                        if (bmp != null)
                        {
                            bmp = PrepareForOcr(bmp);
                            bytes = BitmapSourceToPngBytes(bmp);
                            return true;
                        }
                    }
                    return false;
                }
                catch (COMException)
                {
                    Thread.Sleep(50);
                }
            }

            return false;
        }

        private static bool TryGetClipboardImageFileBytes(out byte[] bytes)
        {
            bytes = Array.Empty<byte>();

            for (int attempt = 0; attempt < 3; attempt++)
            {
                try
                {
                    if (!Clipboard.ContainsFileDropList())
                        return false;

                    var files = Clipboard.GetFileDropList();
                    if (files.Count == 0)
                        return false;

                    string? file = files.Cast<string>().FirstOrDefault(IsLikelyImageFile);

                    if (string.IsNullOrWhiteSpace(file) || !File.Exists(file))
                        return false;

                    bytes = File.ReadAllBytes(file);
                    return true;
                }
                catch (COMException)
                {
                    Thread.Sleep(50);
                }
            }

            return false;
        }

        private void Window_DragOver(object sender, DragEventArgs e)
        {
            e.Effects = HasImageData(e.Data) ? DragDropEffects.Copy : DragDropEffects.None;
            e.Handled = true;
        }

        private async void Window_Drop(object sender, DragEventArgs e)
        {
            if (TryGetImageBytes(e.Data, out var bytes))
            {
                SetImageBytes(bytes);
                await RunOcrAsync();
            }
        }

        private static bool HasImageData(IDataObject data)
        {
            if (data.GetDataPresent(DataFormats.Bitmap))
                return true;

            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = data.GetData(DataFormats.FileDrop) as string[];
                return files != null && files.Any(IsLikelyImageFile);
            }

            return false;
        }

        private static bool TryGetImageBytes(IDataObject data, out byte[] bytes)
        {
            bytes = Array.Empty<byte>();

            if (data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = data.GetData(DataFormats.FileDrop) as string[];
                if (files != null)
                {
                    string? file = files.FirstOrDefault(IsLikelyImageFile);

                    if (!string.IsNullOrWhiteSpace(file) && File.Exists(file))
                    {
                        bytes = File.ReadAllBytes(file);
                        return true;
                    }
                }
            }

            if (data.GetDataPresent(DataFormats.Bitmap))
            {
                var bmp = data.GetData(DataFormats.Bitmap) as BitmapSource;
                if (bmp != null)
                {
                    bytes = BitmapSourceToPngBytes(bmp);
                    return true;
                }
            }

            return false;
        }

        private static bool IsLikelyImageFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            string ext = Path.GetExtension(path);
            return !string.IsNullOrWhiteSpace(ext) && ImageExtensions.Contains(ext);
        }

        private void SetImageBytes(byte[] bytes)
        {
            if (bytes.Length == 0)
                return;

            try
            {
                BitmapImage preview = BytesToBitmapImage(bytes);
                PreviewImage.Source = preview;

                _currentImageBytes = BitmapSourceToPngBytes(preview);
            }
            catch
            {
                PreviewImage.Source = null;
                _currentImageBytes = bytes; // fallback
            }

            OcrTextBox.Clear();
            StatusText.Text = "image loaded.";
            SumResultText.Text = "sum: -";
            UpdateHintVisibility();
        }

        private void UpdateHintVisibility()
        {
            DropHint.Visibility = PreviewImage.Source == null ? Visibility.Visible : Visibility.Collapsed;
        }

        private async void RunOcr_Click(object sender, RoutedEventArgs e)
        {
            await RunOcrAsync();
        }

        private void Sum_Click(object sender, RoutedEventArgs e)
        {
            var (sum, count) = SumNumbersFromText(OcrTextBox.Text);
            SumResultText.Text = $"sum: {sum.ToString("N", CultureInfo.CurrentCulture)}    (numbers found: {count})";
        }

        private async Task RunOcrAsync()
        {
            if (_currentImageBytes.Length == 0)
            {
                StatusText.Text = "no image yet (paste or drop one).";
                return;
            }

            await _ocrLock.WaitAsync();
            try
            {
                StatusText.Text = "running OCR...";
                byte[] bytes = _currentImageBytes;

                string text = await Task.Run(() =>
                {
                    if (!Directory.Exists(TessdataPath))
                        throw new DirectoryNotFoundException("missing tessdata folder: " + TessdataPath);

                    var languages = new List<OcrLanguage>();

                    string eng = Path.Combine(TessdataPath, "eng.traineddata");
                    string deu = Path.Combine(TessdataPath, "deu.traineddata");

                    if (File.Exists(eng)) languages.Add(OcrLanguage.English);
                    if (File.Exists(deu)) languages.Add(OcrLanguage.German);

                    if (languages.Count == 0)
                        throw new FileNotFoundException("no traineddata found. put e.g. traineddata into: " + TessdataPath);

                    var initialValues = new Dictionary<string, object>
                    {
                        ["load_system_dawg"] = 0,
                        ["load_freq_dawg"] = 0,
                        ["user_defined_dpi"] = 30
                    };

                    using (var engine = new Engine(TessdataPath, languages, EngineMode.LstmOnly, initialValues: initialValues))
                    {
                        engine.DefaultPageSegMode = PageSegMode.SparseText;
                        //engine.SetVariable("tessedit_write_images", true);
                        // engine.SetVariable("tessedit_char_whitelist", "0123456789.,-");

                        using (var img = TesseractOCR.Pix.Image.LoadFromMemory(bytes))
                        using (var page = engine.Process(img, PageSegMode.SparseText))
                        {
                            return page.Text ?? string.Empty;
                        }
                    }
                });

                OcrTextBox.Text = NormalizeOcrText(text);
                StatusText.Text = "OCR done.";
            }
            catch (Exception ex)
            {
                StatusText.Text = "OCR error: " + ex.Message;
            }
            finally
            {
                _ocrLock.Release();
            }
        }

        private static (decimal Sum, int Count) SumNumbersFromText(string text)
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

        private static byte[] BitmapSourceToPngBytes(BitmapSource source)
        {
            var encoder = new PngBitmapEncoder();
            encoder.Frames.Add(BitmapFrame.Create(source));

            using (var ms = new MemoryStream())
            {
                encoder.Save(ms);
                return ms.ToArray();
            }
        }

        private static BitmapImage BytesToBitmapImage(byte[] bytes)
        {
            using (var ms = new MemoryStream(bytes))
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.StreamSource = ms;
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
        }

        private static BitmapSource PrepareForOcr(BitmapSource src)
        {
            var rgb = new FormatConvertedBitmap(src, PixelFormats.Bgr24, null, 0);
            rgb.Freeze();

            var scaled = new TransformedBitmap(rgb, new ScaleTransform(2.0, 2.0));
            scaled.Freeze();

            return AddBorder(scaled, 10);
        }

        private static BitmapSource AddBorder(BitmapSource src, int borderPx)
        {
            int w = src.PixelWidth + borderPx * 2;
            int h = src.PixelHeight + borderPx * 2;

            var dv = new DrawingVisual();
            using (var dc = dv.RenderOpen())
            {
                dc.DrawRectangle(Brushes.White, null, new System.Windows.Rect(0, 0, w, h));
                dc.DrawImage(src, new System.Windows.Rect(borderPx, borderPx, src.PixelWidth, src.PixelHeight));
            }

            var rtb = new RenderTargetBitmap(w, h, 96, 96, PixelFormats.Pbgra32);
            rtb.Render(dv);
            rtb.Freeze();
            return rtb;
        }

        private static string NormalizeOcrText(string s)
        {
            if (string.IsNullOrWhiteSpace(s))
                return string.Empty;

            s = s.Replace("\r\n", "\n").Replace('\r', '\n');

            var lines = s.Split('\n')
                        .Select(l => l.Trim())
                        .Where(l => !string.IsNullOrWhiteSpace(l));

            return string.Join(Environment.NewLine, lines).Trim();
        }

        private static bool TryParseByLastSeparatorRule(string raw, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            // normalize minusses
            string s = raw.Trim()
                .Replace('−', '-') // U+2212
                .Replace('–', '-') // U+2013
                .Replace('—', '-') // U+2014
                .Replace('‐', '-'); // U+2010

            // retarded accounting number format for negative numbers
            bool negParen = s.StartsWith("(") && s.EndsWith(")");
            if (negParen)
                s = s.Substring(1, s.Length - 2);

            // remove spaces
            s = s.Replace(" ", "")
                .Replace("\t", "")
                .Replace("\u00A0", "")  // NBSP
                .Replace("\u202F", ""); // narrow NBSP

            // retarded accounting number format for negative numbers part 2
            bool negative = negParen;
            if (s.StartsWith("+"))
                s = s.Substring(1);
            else if (s.StartsWith("-"))
            {
                negative = true;
                s = s.Substring(1);
            }

            // find last decimal seperator
            int lastDot = s.LastIndexOf('.');
            int lastComma = s.LastIndexOf(',');
            int lastSepIndex = Math.Max(lastDot, lastComma);

            int digitsAfterSep = 0;
            if (lastSepIndex >= 0)
                digitsAfterSep = s.Length - lastSepIndex - 1;

            // remove non-digits
            string digitsOnly = new string(s.Where(char.IsDigit).ToArray());
            if (digitsOnly.Length == 0)
                return false;

            // re-insert decimal seperator to create plausible numbers eg ,50 -> 0.50
            string normalized;

            if (lastSepIndex >= 0 && digitsAfterSep > 0 && digitsAfterSep <= MaxDecimalDigits)
            {
                if (digitsOnly.Length <= digitsAfterSep)
                    digitsOnly = digitsOnly.PadLeft(digitsAfterSep + 1, '0');

                normalized = digitsOnly.Insert(digitsOnly.Length - digitsAfterSep, ".");
            }
            else
            {
                normalized = digitsOnly;
            }

            if (!decimal.TryParse(
                    normalized,
                    NumberStyles.AllowDecimalPoint,
                    CultureInfo.InvariantCulture,
                    out var parsed))
                return false;

            value = negative ? -parsed : parsed;
            return true;
        }
    }
}