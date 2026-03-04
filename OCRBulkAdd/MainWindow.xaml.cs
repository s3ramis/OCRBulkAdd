using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
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

        private static readonly Regex NumberTokenRegex = new Regex(
            @"(?:(?:[+\-]|−|–|—)[ \t\u00A0\u202F]*)?(?:(?:\d{1,3}(?:[., \t\u00A0\u202F]\d{3})+|\d+)(?:[.,]\d+)?|[.,]\d+)",
            RegexOptions.Compiled);

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
            if (bytes.Length == 0) return;

            try
            {
                BitmapImage preview = BytesToBitmapImage(bytes);
                PreviewImage.Source = preview;

                var prepared = PrepareForOcr(preview);
                _currentImageBytes = BitmapSourceToPngBytes(prepared);
            }
            catch
            {
                PreviewImage.Source = null;
                _currentImageBytes = bytes;
            }

            OcrPreviewTextBox.Clear();
            NumbersTextBox.Clear();
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
            var (sum, count) = SumNumbersFromText(NumbersTextBox.Text);

            SumResultText.Text = $"sum: {sum.ToString("0.00", CultureInfo.CurrentCulture)} (numbers found: {count})";
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
                        ["user_defined_dpi"] = 300
                    };

                    using (var engine = new Engine(TessdataPath, languages, EngineMode.LstmOnly, initialValues: initialValues))
                    {
                        engine.DefaultPageSegMode = PageSegMode.SparseText;
                        
                        // DEBUG
                        engine.SetVariable("tessedit_write_images", true);

                        using (var img = TesseractOCR.Pix.Image.LoadFromMemory(bytes))
                        using (var page = engine.Process(img, PageSegMode.SparseText))
                        {
                            return page.Text ?? string.Empty;
                        }
                    }
                });

                var previewText = text;

                OcrPreviewTextBox.Text = previewText;


                NumbersTextBox.Text = ExtractNormalizedNumbersText(previewText);

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

        private static string ExtractNormalizedNumbersText(string text)
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

        private static bool TryParseByLastSeparatorRule(string raw, out decimal value)
        {
            value = 0m;
            if (string.IsNullOrWhiteSpace(raw))
                return false;

            // normalize minusses
            string s = raw.Trim()
                .Replace('−', '-')
                .Replace('–', '-')
                .Replace('—', '-')
                .Replace('‐', '-');

            // retarded accounting number format for negative numbers
            bool negParen = s.StartsWith("(") && s.EndsWith(")");
            if (negParen)
                s = s.Substring(1, s.Length - 2);

            // remove spaces
            s = s.Replace(" ", "")
                 .Replace("\t", "")
                 .Replace("\u00A0", "")
                 .Replace("\u202F", "");

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

            // re-insert decimal seperator to assemble a plausible number
            string normalized;

            if (lastSepIndex >= 0 && digitsAfterSep > 0 && digitsAfterSep <= MaxDecimalDigits)
            {
                // padding with a leading 0 if necessary
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

        private static BitmapSource PrepareForOcr(BitmapSource src)
        {
            var opaque = CompositeOnWhite(src);

            var scaled = new TransformedBitmap(opaque, new ScaleTransform(2.0, 2.0));
            scaled.Freeze();

            var gray = new FormatConvertedBitmap(scaled, PixelFormats.Gray8, null, 0);
            gray.Freeze();

            var contrasted = AutoContrastGray8(gray, lowCutPercent: 0.01, highCutPercent: 0.99);

            var binary = BinarizeOtsuGray8(contrasted);

            return AddBorderGray8(binary, borderPx: 12);
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

        private static BitmapSource CompositeOnWhite(BitmapSource src)
        {
            var bgra = new FormatConvertedBitmap(src, PixelFormats.Bgra32, null, 0);
            bgra.Freeze();

            int w = bgra.PixelWidth;
            int h = bgra.PixelHeight;
            int stride = w * 4;

            byte[] px = new byte[stride * h];
            bgra.CopyPixels(px, stride, 0);

            for (int i = 0; i < px.Length; i += 4)
            {
                byte b = px[i];
                byte g = px[i + 1];
                byte r = px[i + 2];
                byte a = px[i + 3];

                if (a != 255)
                {
                    int inv = 255 - a;
                    px[i]     = (byte)((b * a + 255 * inv + 127) / 255);
                    px[i + 1] = (byte)((g * a + 255 * inv + 127) / 255);
                    px[i + 2] = (byte)((r * a + 255 * inv + 127) / 255);
                    px[i + 3] = 255;
                }
                else
                {
                    px[i + 3] = 255;
                }
            }

            var wb = new WriteableBitmap(w, h, 96, 96, PixelFormats.Bgra32, null);
            wb.WritePixels(new Int32Rect(0, 0, w, h), px, stride, 0);
            wb.Freeze();
            return wb;
        }

        private static BitmapSource AutoContrastGray8(BitmapSource gray, double lowCutPercent, double highCutPercent)
        {
            if (gray.Format != PixelFormats.Gray8)
            {
                gray = new FormatConvertedBitmap(gray, PixelFormats.Gray8, null, 0);
                gray.Freeze();
            }

            int w = gray.PixelWidth;
            int h = gray.PixelHeight;
            int stride = w;

            byte[] px = new byte[stride * h];
            gray.CopyPixels(px, stride, 0);

            int[] hist = new int[256];
            for (int i = 0; i < px.Length; i++)
                hist[px[i]]++;

            int total = px.Length;
            int lowTarget = (int)(total * lowCutPercent);
            int highTarget = (int)(total * highCutPercent);

            int cum = 0;
            int low = 0;
            for (int i = 0; i < 256; i++)
            {
                cum += hist[i];
                if (cum >= lowTarget) { low = i; break; }
            }

            cum = 0;
            int high = 255;
            for (int i = 0; i < 256; i++)
            {
                cum += hist[i];
                if (cum >= highTarget) { high = i; break; }
            }

            if (high <= low + 1)
                return gray;

            int range = high - low;

            for (int i = 0; i < px.Length; i++)
            {
                int v = (px[i] - low) * 255 / range;
                if (v < 0) v = 0;
                if (v > 255) v = 255;
                px[i] = (byte)v;
            }

            var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Gray8, null, px, stride);
            bmp.Freeze();
            return bmp;
        }

        private static BitmapSource BinarizeOtsuGray8(BitmapSource gray)
        {
            if (gray.Format != PixelFormats.Gray8)
            {
                gray = new FormatConvertedBitmap(gray, PixelFormats.Gray8, null, 0);
                gray.Freeze();
            }

            int w = gray.PixelWidth;
            int h = gray.PixelHeight;
            int stride = w;

            byte[] px = new byte[stride * h];
            gray.CopyPixels(px, stride, 0);

            int[] hist = new int[256];
            for (int i = 0; i < px.Length; i++)
                hist[px[i]]++;

            int total = px.Length;
            double sum = 0;
            for (int t = 0; t < 256; t++)
                sum += t * hist[t];

            double sumB = 0;
            int wB = 0;
            double maxVar = -1;
            int threshold = 128;

            for (int t = 0; t < 256; t++)
            {
                wB += hist[t];
                if (wB == 0) continue;

                int wF = total - wB;
                if (wF == 0) break;

                sumB += t * hist[t];
                double mB = sumB / wB;
                double mF = (sum - sumB) / wF;

                double varBetween = (double)wB * wF * (mB - mF) * (mB - mF);
                if (varBetween > maxVar)
                {
                    maxVar = varBetween;
                    threshold = t;
                }
            }

            byte[] bin = new byte[px.Length];
            int black = 0, white = 0;

            for (int i = 0; i < px.Length; i++)
            {
                byte v = (px[i] <= threshold) ? (byte)0 : (byte)255;
                bin[i] = v;
                if (v == 0) black++; else white++;
            }

            if (black > white)
            {
                for (int i = 0; i < bin.Length; i++)
                    bin[i] = (byte)(255 - bin[i]);
            }

            var bmp = BitmapSource.Create(w, h, 96, 96, PixelFormats.Gray8, null, bin, stride);
            bmp.Freeze();
            return bmp;
        }

        private static BitmapSource AddBorderGray8(BitmapSource gray, int borderPx)
        {
            if (gray.Format != PixelFormats.Gray8)
            {
                gray = new FormatConvertedBitmap(gray, PixelFormats.Gray8, null, 0);
                gray.Freeze();
            }

            int w = gray.PixelWidth;
            int h = gray.PixelHeight;
            int stride = w;

            byte[] srcPx = new byte[stride * h];
            gray.CopyPixels(srcPx, stride, 0);

            int newW = w + borderPx * 2;
            int newH = h + borderPx * 2;
            int newStride = newW;

            byte[] outPx = new byte[newStride * newH];
            Array.Fill(outPx, (byte)255);

            for (int y = 0; y < h; y++)
            {
                Buffer.BlockCopy(
                    srcPx, y * stride,
                    outPx, (y + borderPx) * newStride + borderPx,
                    w);
            }

            var bmp = BitmapSource.Create(newW, newH, 96, 96, PixelFormats.Gray8, null, outPx, newStride);
            bmp.Freeze();
            return bmp;
        }
    }
}